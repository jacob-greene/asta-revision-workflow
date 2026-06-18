#!/usr/bin/env python3
"""Flag modified DOCX sentences that lack an adjacent citation.

This is a workflow guardrail, not a semantic proof checker. If it flags a
modified sentence, the revision agent must either verify that nearby existing
citations support the statement, add a citation, soften/delete the claim, or
query Asta for targeted evidence.
"""

from __future__ import annotations

import argparse
import re
import zipfile
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from xml.etree import ElementTree as ET

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"

CITE_TOKEN_RE = re.compile(r"(?:\[\[CITE:[^\]]+\]\]|\{[^{}]+,\s*\d{4}[^{}]*\})")
REFERENCE_START_RE = re.compile(r"^\s*1\.\s*[A-Z]")
SENTENCE_BOUNDARY_RE = re.compile(r"(?<=[.!?])\s+(?=[A-Z])")


@dataclass(frozen=True)
class Sentence:
    text: str
    comparable: str
    has_citation: bool


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def comparable(text: str) -> str:
    text = CITE_TOKEN_RE.sub("", text)
    text = re.sub(r"\b\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*\b", "", text)
    text = re.sub(r"[^A-Za-z0-9]+", " ", text.lower())
    return clean(text)


def text_of(elem: ET.Element) -> str:
    return "".join(t.text or "" for t in elem.findall(f".//{W}t"))


def is_superscript_run(run: ET.Element) -> bool:
    rpr = run.find(f"{W}rPr")
    if rpr is None:
        return False
    vert = rpr.find(f"{W}vertAlign")
    return vert is not None and vert.attrib.get(f"{W}val") == "superscript"


def paragraph_text(paragraph: ET.Element) -> str:
    pieces: list[str] = []
    for run in paragraph.findall(f"{W}r"):
        run_text = text_of(run)
        if is_superscript_run(run) and re.fullmatch(r"\s*\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*\s*", run_text):
            pieces.append(f" [[CITE:{clean(run_text)}]]")
        else:
            pieces.append(run_text)
    return clean("".join(pieces))


def docx_paragraphs(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(f"{W}body")
    if body is None:
        raise RuntimeError("word/document.xml has no body")
    paragraphs = [paragraph_text(p) for p in body.findall(f"{W}p")]
    return [p for p in paragraphs if p]


def content_paragraphs(path: Path) -> list[str]:
    paragraphs = docx_paragraphs(path)
    for index, paragraph in enumerate(paragraphs):
        if REFERENCE_START_RE.match(paragraph):
            return paragraphs[:index]
    return paragraphs


def split_sentences(paragraph: str) -> list[Sentence]:
    sentences: list[Sentence] = []
    protected = paragraph
    placeholders = {
        "H3.1": "H3<dot>1",
        "H3.2": "H3<dot>2",
        "H3.3": "H3<dot>3",
        "e.g.": "e<dot>g<dot>",
        "i.e.": "i<dot>e<dot>",
    }
    for original, replacement in placeholders.items():
        protected = protected.replace(original, replacement)
    for chunk in SENTENCE_BOUNDARY_RE.split(protected):
        text = clean(chunk)
        if not text:
            continue
        for original, replacement in placeholders.items():
            text = text.replace(replacement, original)
        sentences.append(Sentence(text=text, comparable=comparable(text), has_citation=bool(CITE_TOKEN_RE.search(text))))
    return sentences


def sentence_has_adjacent_citation(sentences: list[Sentence], index: int) -> bool:
    for neighbor in (index - 1, index, index + 1):
        if 0 <= neighbor < len(sentences) and sentences[neighbor].has_citation:
            return True
    return False


def parse_indices(value: str | None) -> set[int] | None:
    if not value:
        return None
    indices: set[int] = set()
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            indices.update(range(int(start), int(end) + 1))
        else:
            indices.add(int(part))
    return indices


def aligned_revised_paragraphs(source: list[str], revised: list[str]) -> dict[int, str]:
    """Map 1-based source paragraph indices to their revised paragraph text.

    The launcher permits comment-scoped paragraph splits. A positional zip would
    compare every downstream paragraph to the wrong source paragraph after such
    a split, so align on unchanged paragraph text and join split replacements.
    """

    mapping: dict[int, str] = {}
    matcher = SequenceMatcher(a=source, b=revised, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for offset in range(i2 - i1):
                mapping[i1 + offset + 1] = revised[j1 + offset]
            continue
        if tag == "delete":
            continue
        if tag == "insert":
            continue
        if tag == "replace":
            source_span = i2 - i1
            revised_span = j2 - j1
            if source_span == revised_span:
                for offset in range(source_span):
                    mapping[i1 + offset + 1] = revised[j1 + offset]
            elif source_span == 1:
                mapping[i1 + 1] = clean(" ".join(revised[j1:j2]))
            else:
                for offset in range(source_span):
                    raw_index = min(j1 + offset, j2 - 1)
                    mapping[i1 + offset + 1] = revised[raw_index] if j1 < j2 else ""
    return mapping


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, help="Original commented/source DOCX.")
    parser.add_argument("--revised", required=True, help="Revised raw DOCX before reference-list removal.")
    parser.add_argument(
        "--paragraphs",
        help="Optional comma/range list of 1-based content paragraph indices to check, e.g. 8,9,12-16.",
    )
    args = parser.parse_args()

    source = content_paragraphs(Path(args.source))
    revised = content_paragraphs(Path(args.revised))
    revised_by_source = aligned_revised_paragraphs(source, revised)
    selected = parse_indices(args.paragraphs)
    failures: list[tuple[int, str]] = []

    for paragraph_index in range(1, len(source) + 1):
        if selected is not None and paragraph_index not in selected:
            continue
        if paragraph_index not in revised_by_source:
            continue
        old = comparable(source[paragraph_index - 1])
        revised_text = revised_by_source[paragraph_index]
        new = comparable(revised_text)
        if old == new:
            continue
        old_sentences = {sentence.comparable for sentence in split_sentences(source[paragraph_index - 1])}
        new_sentences = split_sentences(revised_text)
        for sentence_index, sentence in enumerate(new_sentences):
            if not sentence.comparable or sentence.comparable in old_sentences:
                continue
            if not sentence_has_adjacent_citation(new_sentences, sentence_index):
                failures.append((paragraph_index, sentence.text))

    if failures:
        print("FAIL: modified sentences without same/adjacent citation.")
        for paragraph_index, sentence in failures[:50]:
            print(f"Paragraph {paragraph_index}: {sentence}")
        print(
            "Action required: verify nearby existing citations, add/adjust citations, soften/delete unsupported claims, "
            "or requery Asta for targeted evidence before delivery."
        )
        raise SystemExit(1)

    print("PASS: modified sentences have same/adjacent citation support.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
