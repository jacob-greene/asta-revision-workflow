#!/usr/bin/env python3
"""Convert numeric superscript citations in a DOCX to EndNote temporary citations."""

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
ET.register_namespace("w", W_URI)

CITE_RE = re.compile(r"^\s*\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*\s*$")
REF_START_RE = re.compile(r"(?<!\d)(\d{1,2})\.\s*(?=[A-Z])")


@dataclass(frozen=True)
class Reference:
    author: str
    year: str
    title: str

    @property
    def author_year(self) -> str:
        return f"{self.author}, {self.year}"


def text_of(elem: ET.Element) -> str:
    return "".join(t.text or "" for t in elem.findall(f".//{W}t"))


def is_superscript_run(run: ET.Element) -> bool:
    rpr = run.find(f"{W}rPr")
    if rpr is None:
        return False
    vert = rpr.find(f"{W}vertAlign")
    return vert is not None and vert.attrib.get(f"{W}val") == "superscript"


def parse_references(reference_text: str) -> dict[int, Reference]:
    mapping: dict[int, Reference] = {}
    normalized = re.sub(r"\s+", " ", reference_text.strip())
    starts = list(REF_START_RE.finditer(normalized))
    for i, match in enumerate(starts):
        number = int(match.group(1))
        end = starts[i + 1].start() if i + 1 < len(starts) else len(normalized)
        entry = normalized[match.end() : end].strip()
        author = first_author(entry)
        year = first_year(entry)
        title = title_from_entry(entry)
        if author and year:
            mapping[number] = Reference(author=author, year=year, title=title)
    return mapping


def first_author(entry: str) -> str:
    before_year = re.split(r"\(\d{4}\)", entry, maxsplit=1)[0]
    first = before_year.split(",", 1)[0].strip()
    return re.sub(r"\s+", " ", first)


def first_year(entry: str) -> str:
    match = re.search(r"\((\d{4})\)", entry)
    return match.group(1) if match else ""


def title_from_entry(entry: str) -> str:
    after_year = re.split(r"\(\d{4}\)\.", entry, maxsplit=1)
    if len(after_year) < 2:
        return ""
    title = after_year[1].split(".", 1)[0].strip()
    return re.sub(r"\s+", " ", title)


def normalized_title(title: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", title.lower()).strip()


def temporary_citation_text(reference: Reference, ambiguous: bool) -> str:
    if ambiguous and reference.title:
        return f"{reference.author}, {reference.year}, {reference.title}"
    return reference.author_year


def expand_citation_numbers(text: str) -> list[int]:
    numbers: list[int] = []
    for part in text.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            numbers.extend(range(int(start), int(end) + 1))
        else:
            numbers.append(int(part))
    return numbers


def make_temp_citation(text: str, references: dict[int, Reference], ambiguous_keys: set[str]) -> str:
    parts: list[str] = []
    for number in expand_citation_numbers(text):
        reference = references.get(number)
        if reference is None:
            parts.append(f"REF{number}")
        else:
            parts.append(temporary_citation_text(reference, reference.author_year in ambiguous_keys))
    return "{" + "; ".join(parts) + "}"


def set_run_text(run: ET.Element, text: str) -> None:
    rpr = run.find(f"{W}rPr")
    if rpr is not None:
        for child in list(rpr):
            if child.tag == f"{W}vertAlign":
                rpr.remove(child)
        if len(list(rpr)) == 0 and not rpr.attrib:
            run.remove(rpr)
    for child in list(run):
        if child.tag != f"{W}rPr":
            run.remove(child)
    t = ET.SubElement(run, f"{W}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def convert(source: Path, output: Path, keep_references: bool = False) -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(source) as zf:
            zf.extractall(tmp)

        document = tmp / "word" / "document.xml"
        tree = ET.parse(document)
        root = tree.getroot()
        body = root.find(f"{W}body")
        if body is None:
            raise RuntimeError("word/document.xml has no body")

        paragraphs = [p for p in body.findall(f"{W}p") if text_of(p).strip()]
        ref_index = None
        for i, paragraph in enumerate(paragraphs):
            if re.match(r"^\s*1\.\s*[A-Z]", text_of(paragraph).strip()):
                ref_index = i
                break
        if ref_index is None:
            raise RuntimeError("Could not locate numbered reference list in DOCX.")

        reference_text = " ".join(text_of(p) for p in paragraphs[ref_index:])
        references = parse_references(reference_text)
        if not references:
            raise RuntimeError("Could not parse numbered references.")
        author_year_titles: dict[str, set[str]] = {}
        for reference in references.values():
            author_year_titles.setdefault(reference.author_year, set()).add(normalized_title(reference.title))
        ambiguous_keys = {key for key, titles in author_year_titles.items() if len(titles) > 1}

        for paragraph in paragraphs[:ref_index]:
            preceding = ""
            for run in paragraph.findall(f"{W}r"):
                run_text = text_of(run)
                if is_superscript_run(run) and CITE_RE.match(run_text):
                    citation = make_temp_citation(run_text, references, ambiguous_keys)
                    if preceding and not preceding[-1].isspace():
                        citation = " " + citation
                    set_run_text(run, citation)
                    preceding += citation
                else:
                    preceding += run_text

        if not keep_references:
            for paragraph in paragraphs[ref_index:]:
                parent_children = list(body)
                if paragraph in parent_children:
                    body.remove(paragraph)

        tree.write(document, encoding="UTF-8", xml_declaration=True)
        if output.exists():
            output.unlink()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in sorted(tmp.rglob("*")):
                if path.is_file():
                    zf.write(path, path.relative_to(tmp).as_posix())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source")
    parser.add_argument("output")
    parser.add_argument("--keep-references", action="store_true")
    args = parser.parse_args()
    convert(Path(args.source), Path(args.output), keep_references=args.keep_references)
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
