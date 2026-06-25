#!/usr/bin/env python3
"""Word-focused sanity checks for generated DOCX files."""

from __future__ import annotations

import argparse
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

W14_URI = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_URI = "http://schemas.microsoft.com/office/word/2012/wordml"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = f"{{{W14_URI}}}"
W15 = f"{{{W15_URI}}}"


def attr_values(root: ET.Element, attr_name: str) -> list[str]:
    values: list[str] = []
    for elem in root.iter():
        for attr, value in elem.attrib.items():
            if attr.split("}")[-1] == attr_name:
                values.append(value)
    return values


def duplicate_values(values: list[str]) -> list[str]:
    counts = Counter(values)
    return sorted(value for value, count in counts.items() if count > 1)


def check_docx(path: Path) -> list[str]:
    errors: list[str] = []
    with zipfile.ZipFile(path) as zf:
        document = ET.fromstring(zf.read("word/document.xml"))
        document_text = ET.tostring(document, encoding="unicode")
        for token in ["commentRangeStart", "commentRangeEnd", "commentReference"]:
            if token in document_text:
                errors.append(f"stale Word comment marker remains: {token}")
        for attr_name in ["paraId", "textId"]:
            values = attr_values(document, attr_name)
            if attr_name == "textId":
                values = [value for value in values if value != "77777777"]
            duplicates = duplicate_values(values)
            if duplicates:
                preview = ", ".join(duplicates[:5])
                errors.append(f"duplicate {attr_name} values: {preview}")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx")
    args = parser.parse_args()
    errors = check_docx(Path(args.docx))
    if errors:
        for error in errors:
            print(f"FAIL: {error}")
        return 1
    print("PASS: Word sanity checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
