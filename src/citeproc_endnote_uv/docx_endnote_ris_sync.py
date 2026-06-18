#!/usr/bin/env python3
"""Check that DOCX EndNote temporary citations resolve to a paired RIS file."""

from __future__ import annotations

import argparse
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"


@dataclass(frozen=True)
class RisRecord:
    authors: tuple[str, ...]
    year: str
    title: str

    @property
    def author_year(self) -> str:
        return f"{self.authors[0].split(',', 1)[0]}, {self.year}"


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def normalized_title(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def records_from_ris(path: Path) -> list[RisRecord]:
    records: list[RisRecord] = []
    current: dict[str, str] = {}
    authors: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith("ER  -"):
            if authors and current.get("PY"):
                records.append(
                    RisRecord(
                        authors=tuple(authors),
                        year=clean(current.get("PY", "")),
                        title=clean(current.get("TI", "")),
                    )
                )
            current = {}
            authors = []
            continue
        match = re.match(r"([A-Z0-9]{2})  - (.*)", line)
        if not match:
            continue
        key, value = match.group(1), clean(match.group(2))
        if key == "AU":
            authors.append(value)
        else:
            current[key] = value
    return records


def docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    return "".join(t.text or "" for t in root.findall(f".//{W}t"))


def temporary_entries(text: str) -> list[str]:
    entries: list[str] = []
    for citation in re.findall(r"\{([^{}]+)\}", text):
        entries.extend(clean(part) for part in citation.split(";") if clean(part))
    return entries


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", help="DOCX containing EndNote temporary citations.")
    parser.add_argument("ris", help="Same-stem RIS file generated from the revised reference list.")
    parser.add_argument(
        "--allow-ambiguous-author-year",
        action="store_true",
        help="Allow {Author, Year} temporary citations that match multiple RIS records. Use only for EndNote-safe documents that avoid long title disambiguation.",
    )
    parser.add_argument(
        "--allow-title-prefix",
        action="store_true",
        help="Allow abbreviated title prefixes in ambiguous temporary citations.",
    )
    args = parser.parse_args()

    records = records_from_ris(Path(args.ris))
    by_author_year: dict[str, list[RisRecord]] = {}
    by_author_year_title: dict[tuple[str, str], RisRecord] = {}
    for record in records:
        by_author_year.setdefault(record.author_year, []).append(record)
        if record.title:
            by_author_year_title[(record.author_year, record.title.lower())] = record

    entries = temporary_entries(docx_text(Path(args.docx)))
    missing: list[str] = []
    ambiguous_without_title: list[str] = []
    ref_placeholders: list[str] = []
    matched = 0

    for entry in entries:
        if re.fullmatch(r"REF\d+", entry):
            ref_placeholders.append(entry)
            continue
        pieces = [clean(piece) for piece in entry.split(",")]
        if len(pieces) < 2:
            missing.append(entry)
            continue
        author_year = f"{pieces[0]}, {pieces[1]}"
        title = clean(",".join(pieces[2:]))
        candidates = by_author_year.get(author_year, [])
        if not candidates:
            missing.append(entry)
            continue
        if len(candidates) > 1:
            if not title:
                if args.allow_ambiguous_author_year:
                    matched += 1
                    continue
                else:
                    ambiguous_without_title.append(entry)
                    continue
            if (author_year, title.lower()) not in by_author_year_title:
                if args.allow_title_prefix:
                    normalized_prefix = normalized_title(title)
                    prefix_matches = [
                        record
                        for record in candidates
                        if normalized_title(record.title).startswith(normalized_prefix)
                    ]
                    matched_titles = {normalized_title(record.title) for record in prefix_matches}
                    if len(prefix_matches) == 1 or len(matched_titles) == 1:
                        matched += 1
                        continue
                missing.append(entry)
                continue
        matched += 1

    print(f"RIS records: {len(records)}")
    print(f"Temporary citation entries: {len(entries)}")
    print(f"Matched entries: {matched}")
    print(f"REF placeholders: {len(ref_placeholders)}")
    print(f"Missing entries: {len(missing)}")
    print(f"Ambiguous entries without title: {len(ambiguous_without_title)}")

    for label, values in (
        ("REF", ref_placeholders),
        ("MISSING", missing),
        ("AMBIGUOUS", ambiguous_without_title),
    ):
        for value in values[:25]:
            print(f"{label}: {value}")

    if args.allow_ambiguous_author_year and ambiguous_without_title:
        print("WARNING: ambiguous author-year citations were allowed.")

    if ref_placeholders or missing or ambiguous_without_title:
        raise SystemExit(1)
    print("PASS: EndNote temporary citations match the RIS.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
