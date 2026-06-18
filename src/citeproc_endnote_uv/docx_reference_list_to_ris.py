#!/usr/bin/env python3
"""Export a numbered Word reference list from a DOCX package to RIS.

This is intentionally driven by the rebuilt DOCX rather than by a separate
BibTeX file, so references manually added to the Word source propagate into the
paired EndNote import file for the same pass.
"""

from __future__ import annotations

import argparse
import re
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"

REF_START_RE = re.compile(r"(?:^|\n)\s*(\d{1,3})\.\s*(?=[A-Z])")
YEAR_RE = re.compile(r"\((\d{4})\)")
DOI_RE = re.compile(r"\b10\.\d{4,9}/\S+")
PMID_RE = re.compile(r"\bPMID:\s*(\d+)", re.IGNORECASE)
DOI_URL_RE = re.compile(r"https?://(?:dx\.)?doi\.org/\S*", re.IGNORECASE)


@dataclass(frozen=True)
class Reference:
    number: int
    authors: str
    year: str
    title: str
    journal: str
    volume: str
    pages: str
    doi: str
    pmid: str

    @property
    def key(self) -> tuple[str, str, str]:
        first_author = self.authors.split(",", 1)[0].lower()
        title = re.sub(r"[^a-z0-9]+", " ", self.title.lower()).strip()
        return (first_author, self.year, title)


def text_of(elem: ET.Element) -> str:
    return "".join(t.text or "" for t in elem.findall(f".//{W}t"))


def clean(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def field_key(reference: Reference) -> str:
    first = re.sub(r"[^A-Za-z0-9]+", "", reference.authors.split(",", 1)[0])
    words = re.findall(r"[A-Za-z0-9]+", reference.title)
    suffix = "".join(word[:1].upper() + word[1:8] for word in words[:3])
    return f"{first}{reference.year}{suffix}" if first and reference.year else f"ref{reference.number}"


def reference_text_from_docx(docx: Path) -> str:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(docx) as zf:
            zf.extractall(tmp)
        tree = ET.parse(tmp / "word" / "document.xml")
        body = tree.getroot().find(f"{W}body")
        if body is None:
            raise RuntimeError("word/document.xml has no body")
        paragraphs = [p for p in body.findall(f"{W}p") if text_of(p).strip()]
        ref_index = None
        for index, paragraph in enumerate(paragraphs):
            if re.match(r"^\s*1\.\s*[A-Z]", text_of(paragraph).strip()):
                ref_index = index
                break
        if ref_index is None:
            raise RuntimeError("Could not locate numbered reference list in DOCX.")
        reference_paragraphs: list[str] = []
        for paragraph in paragraphs[ref_index:]:
            text = clean(text_of(paragraph))
            if reference_paragraphs and not REF_START_RE.match(text):
                break
            reference_paragraphs.append(text)
        return "\n".join(reference_paragraphs)


def split_reference_entries(reference_text: str) -> list[tuple[int, str]]:
    starts = list(REF_START_RE.finditer(reference_text))
    entries: list[tuple[int, str]] = []
    for index, match in enumerate(starts):
        number = int(match.group(1))
        end = starts[index + 1].start() if index + 1 < len(starts) else len(reference_text)
        entries.append((number, clean(reference_text[match.end() : end])))
    return entries


def first_sentence_after_year(entry: str) -> tuple[str, str]:
    split = re.split(r"\(\d{4}\)\.", entry, maxsplit=1)
    if len(split) < 2:
        return "", ""
    rest = split[1].strip()
    if "." not in rest:
        return clean(rest), ""
    title, tail = rest.split(".", 1)
    return clean(title), clean(tail)


def split_end_year_entry(entry: str) -> tuple[str, str, str] | None:
    """Parse references that place the year at the end, e.g. Nature 492, 108-112 (2012)."""

    year_match = re.search(r"\((\d{4})\)\.?$", entry)
    if not year_match:
        return None
    before_year = clean(entry[: year_match.start()])
    if ". " not in before_year:
        return None

    if " et al. " in before_year:
        authors, rest = before_year.split(" et al. ", 1)
        authors = clean(f"{authors} et al.")
    else:
        authors = ""
        rest = ""
        for match in re.finditer(r"\.\s+", before_year):
            candidate_rest = before_year[match.end() :]
            if candidate_rest.startswith("& ") or candidate_rest.lower().startswith("and "):
                continue
            if candidate_rest[:1].isupper():
                authors = clean(before_year[: match.end() - 1])
                rest = clean(candidate_rest)
                break
        if not authors:
            return None
    if ". " in rest:
        title, tail = rest.split(". ", 1)
    else:
        title, tail = rest.rstrip("."), ""
    return authors, clean(title), clean(tail)


def parse_tail(tail: str) -> tuple[str, str, str]:
    tail = DOI_RE.sub("", tail)
    tail = DOI_URL_RE.sub("", tail)
    tail = PMID_RE.sub("", tail)
    tail = clean(tail.strip(" ."))
    match = re.match(r"(?P<journal>.+?)\s+(?P<volume>\d+[A-Za-z]?)\s*,\s*(?P<pages>[A-Za-z]?\d+[^.]*)$", tail)
    if not match:
        return tail, "", ""
    return clean(match.group("journal")), clean(match.group("volume")), clean(match.group("pages"))


def parse_reference(number: int, entry: str) -> Reference | None:
    end_year = split_end_year_entry(entry)
    if end_year is not None:
        authors, title, tail = end_year
        year = re.search(r"\((\d{4})\)\.?$", entry).group(1)  # type: ignore[union-attr]
        journal, volume, pages = parse_tail(tail)
        doi_match = DOI_RE.search(entry)
        pmid_match = PMID_RE.search(entry)
        doi = doi_match.group(0).rstrip(".") if doi_match else ""
        pmid = pmid_match.group(1) if pmid_match else ""
        return Reference(number, authors, year, title, journal, volume, pages, doi, pmid)

    year_match = YEAR_RE.search(entry)
    if not year_match:
        return None
    authors = clean(entry[: year_match.start()].strip())
    year = year_match.group(1)
    title, tail = first_sentence_after_year(entry)
    journal, volume, pages = parse_tail(tail)
    doi_match = DOI_RE.search(entry)
    pmid_match = PMID_RE.search(entry)
    doi = doi_match.group(0).rstrip(".") if doi_match else ""
    pmid = pmid_match.group(1) if pmid_match else ""
    return Reference(number, authors, year, title, journal, volume, pages, doi, pmid)


def author_list(authors: str) -> list[str]:
    authors = clean(authors)
    if not authors:
        return []
    authors = re.sub(r"\bet\s+al\.?$", "", authors, flags=re.IGNORECASE).strip(" ,")
    authors = re.sub(r"\band\s+others$", "", authors, flags=re.IGNORECASE).strip(" ,")
    authors = re.sub(r"\s*&\s*", ", ", authors)
    authors = re.sub(r",?\s+and\s+", ", ", authors, flags=re.IGNORECASE)
    chunks = [chunk.strip() for chunk in authors.split(",")]
    parsed: list[str] = []
    index = 0
    while index + 1 < len(chunks):
        surname = re.sub(r"^(?:and\s+)+", "", chunks[index].strip(), flags=re.IGNORECASE).strip()
        initials = chunks[index + 1].strip()
        if surname and initials and not re.fullmatch(r"(?:and\s+)?(?:others|et al\.?)", surname, re.IGNORECASE):
            parsed.append(f"{surname}, {initials}")
        index += 2
    if not parsed and authors:
        parsed.append(authors)
    return parsed


def write_record(reference: Reference) -> list[str]:
    lines = ["TY  - JOUR"]
    if reference.title:
        lines.append(f"TI  - {reference.title}")
    for author in author_list(reference.authors):
        lines.append(f"AU  - {author}")
    if reference.year:
        lines.append(f"PY  - {reference.year}")
    if reference.journal:
        lines.append(f"JO  - {reference.journal}")
    if reference.volume:
        lines.append(f"VL  - {reference.volume}")
    if reference.pages:
        pages = reference.pages.replace("--", "-")
        if "-" in pages:
            start, end = pages.split("-", 1)
            lines.append(f"SP  - {start}")
            lines.append(f"EP  - {end}")
        else:
            lines.append(f"SP  - {pages}")
    if reference.doi:
        lines.append(f"DO  - {reference.doi}")
    if reference.pmid:
        lines.append(f"AN  - {reference.pmid}")
    lines.append(f"ID  - {field_key(reference)}")
    lines.append("ER  -")
    return lines


def export_ris(source_docx: Path, output_ris: Path) -> int:
    references: list[Reference] = []
    failed: list[tuple[int, str]] = []
    for number, entry in split_reference_entries(reference_text_from_docx(source_docx)):
        reference = parse_reference(number, entry)
        if reference is None:
            failed.append((number, entry[:220]))
            continue
        references.append(reference)

    if failed:
        details = "\n".join(f"{number}. {snippet}" for number, snippet in failed[:20])
        raise RuntimeError(f"Could not parse {len(failed)} reference entries:\n{details}")

    records: list[str] = []
    for reference in references:
        records.extend(write_record(reference))
        records.append("")
    output_ris.write_text("\n".join(records), encoding="utf-8")
    return len(references)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source_docx")
    parser.add_argument("output_ris")
    args = parser.parse_args()
    count = export_ris(Path(args.source_docx), Path(args.output_ris))
    print(f"Wrote {count} records to {args.output_ris}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
