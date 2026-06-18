#!/usr/bin/env python3
"""Extract embedded EndNote records from a DOCX into RIS.

Formatted EndNote Word documents often contain complete reference metadata in
field-code XML even when the visible bibliography is abbreviated with "et al.".
This exporter uses those embedded records as the deterministic metadata source.
"""

from __future__ import annotations

import argparse
import base64
import html
import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"
NS = {"w": W_URI}


@dataclass(frozen=True)
class EndNoteRecord:
    rec_number: str
    title: str
    authors: tuple[str, ...]
    year: str
    journal: str
    volume: str
    issue: str
    pages: str
    doi: str
    pmid: str


def clean(text: str | None) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def normalized_title(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def field_key(record: EndNoteRecord) -> str:
    first = re.sub(r"[^A-Za-z0-9]+", "", record.authors[0].split(",", 1)[0] if record.authors else "")
    words = re.findall(r"[A-Za-z0-9]+", record.title)
    suffix = "".join(word[:1].upper() + word[1:8] for word in words[:3])
    return f"{first}{record.year}{suffix}" if first and record.year else f"rec{record.rec_number or suffix or 'unknown'}"


def text_or_empty(elem: ET.Element | None) -> str:
    return clean("".join(elem.itertext()) if elem is not None else "")


def first_text(record: ET.Element, path: str) -> str:
    return text_or_empty(record.find(path))


def record_from_element(record: ET.Element) -> EndNoteRecord | None:
    title = first_text(record, "titles/title")
    year = first_text(record, "dates/year")
    authors = tuple(
        clean(author.text)
        for author in record.findall("contributors/authors/author")
        if clean(author.text)
    )
    if not title or not year:
        return None
    journal = first_text(record, "titles/secondary-title") or first_text(record, "periodical/full-title")
    doi = first_text(record, "electronic-resource-num")
    pmid = first_text(record, "accession-num")
    return EndNoteRecord(
        rec_number=first_text(record, "rec-number"),
        title=title,
        authors=authors,
        year=year,
        journal=journal,
        volume=first_text(record, "volume"),
        issue=first_text(record, "number"),
        pages=first_text(record, "pages"),
        doi=doi,
        pmid=pmid if pmid.isdigit() else "",
    )


def xml_payloads_from_docx(docx: Path) -> list[str]:
    with zipfile.ZipFile(docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))

    payloads: list[str] = []
    for node in root.findall(".//w:instrText", NS):
        text = node.text or ""
        if "<record>" in text or "&lt;record&gt;" in text:
            payloads.append(html.unescape(text))

    for node in root.findall(".//w:fldData", NS):
        encoded = re.sub(r"\s+", "", node.text or "")
        if not encoded:
            continue
        try:
            decoded = base64.b64decode(encoded + "=" * (-len(encoded) % 4), validate=False).decode("utf-8", "ignore")
        except Exception:
            continue
        if "<record>" in decoded or "&lt;record&gt;" in decoded:
            payloads.append(html.unescape(decoded))
    return payloads


def records_from_payload(payload: str) -> list[EndNoteRecord]:
    records: list[EndNoteRecord] = []
    for match in re.finditer(r"<record>.*?</record>", payload, flags=re.DOTALL):
        record_xml = sanitize_record_xml(match.group(0))
        try:
            elem = ET.fromstring(record_xml)
        except ET.ParseError:
            continue
        record = record_from_element(elem)
        if record is not None:
            records.append(record)
    return records


def sanitize_record_xml(record_xml: str) -> str:
    record_xml = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", record_xml)
    return re.sub(r"&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9A-Fa-f]+);)", "&amp;", record_xml)


def score(record: EndNoteRecord) -> tuple[int, int, int, int, int]:
    return (
        len(record.authors),
        int(bool(record.doi)),
        int(bool(record.journal)),
        int(bool(record.volume)),
        len(record.pages),
    )


def dedupe_records(records: list[EndNoteRecord]) -> list[EndNoteRecord]:
    best: dict[tuple[str, str], EndNoteRecord] = {}
    order: list[tuple[str, str]] = []
    for record in records:
        key = (normalized_title(record.title), record.year)
        if not key[0]:
            continue
        if key not in best:
            best[key] = record
            order.append(key)
            continue
        if score(record) > score(best[key]):
            best[key] = record
    return [best[key] for key in order]


def records_from_docx(docx: Path) -> list[EndNoteRecord]:
    records: list[EndNoteRecord] = []
    for payload in xml_payloads_from_docx(docx):
        records.extend(records_from_payload(payload))
    return dedupe_records(records)


def ris_lines(record: EndNoteRecord, key: str) -> list[str]:
    lines = ["TY  - JOUR", f"TI  - {record.title}"]
    for author in record.authors:
        lines.append(f"AU  - {author}")
    lines.append(f"PY  - {record.year}")
    if record.journal:
        lines.append(f"JO  - {record.journal}")
    if record.volume:
        lines.append(f"VL  - {record.volume}")
    if record.issue:
        lines.append(f"IS  - {record.issue}")
    if record.pages:
        if "-" in record.pages:
            start, end = record.pages.split("-", 1)
            lines.extend([f"SP  - {start}", f"EP  - {end}"])
        else:
            lines.append(f"SP  - {record.pages}")
    if record.doi:
        lines.append(f"DO  - {record.doi}")
    if record.pmid:
        lines.append(f"AN  - {record.pmid}")
    if record.rec_number:
        lines.append(f"N1  - EndNote RecNum {record.rec_number}")
    lines.append(f"ID  - {key}")
    lines.append("ER  -")
    return lines


def ris_text(records: list[EndNoteRecord]) -> str:
    key_counts: dict[str, int] = {}
    for record in records:
        key = field_key(record)
        key_counts[key] = key_counts.get(key, 0) + 1
    chunks: list[str] = []
    for record in records:
        key = field_key(record)
        if key_counts[key] > 1:
            key = f"{key}Rec{record.rec_number or len(chunks) + 1}"
        chunks.append("\n".join(ris_lines(record, key)))
    return "\n\n".join(chunks) + ("\n" if chunks else "")


def export_ris(docx: Path, output_ris: Path) -> dict[str, object]:
    records = records_from_docx(docx)
    output_ris.write_text(ris_text(records), encoding="utf-8")
    missing_authors = [
        {"title": record.title, "year": record.year}
        for record in records
        if not record.authors
    ]
    return {
        "source_docx": str(docx),
        "output_ris": str(output_ris),
        "embedded_record_count": len(records),
        "missing_author_records": missing_authors,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source_docx")
    parser.add_argument("output_ris")
    parser.add_argument("--audit", help="Optional JSON audit path.")
    args = parser.parse_args()

    audit = export_ris(Path(args.source_docx), Path(args.output_ris))
    if args.audit:
        Path(args.audit).write_text(json.dumps(audit, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"Wrote {audit['embedded_record_count']} embedded EndNote records to {args.output_ris}")
    if audit["missing_author_records"]:
        raise SystemExit(f"Embedded EndNote records missing authors: {audit['missing_author_records'][:5]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
