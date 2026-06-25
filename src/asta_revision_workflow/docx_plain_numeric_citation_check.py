#!/usr/bin/env python3
"""Fail if likely numeric citations remain as plain DOCX text."""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"

REF_START_RE = re.compile(r"^\s*(\d{1,3})\.\s*(?=[A-Z])")
NUMERIC_CITATION_RE = re.compile(r"\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*")
TOKEN_ENDING_WITH_BIO_NUMBER_RE = re.compile(
    r"(?:PRC\d+|H\d(?:\.\d)?K27(?:me\d)?|H\d(?:\.\d)?K27M|H\d(?:\.\d)?)$"
)


def text_of(elem: ET.Element) -> str:
    return "".join(t.text or "" for t in elem.findall(f".//{W}t"))


def is_superscript_run(run: ET.Element) -> bool:
    rpr = run.find(f"{W}rPr")
    if rpr is None:
        return False
    vert = rpr.find(f"{W}vertAlign")
    return vert is not None and vert.attrib.get(f"{W}val") == "superscript"


def expand_citation_numbers(text: str) -> list[int]:
    numbers: list[int] = []
    for part in text.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            if not start.isdigit() or not end.isdigit():
                return []
            start_i, end_i = int(start), int(end)
            if start_i > end_i:
                return []
            numbers.extend(range(start_i, end_i + 1))
        elif part.isdigit():
            numbers.append(int(part))
        else:
            return []
    return numbers


def is_valid_citation(text: str, valid_refs: set[int]) -> bool:
    numbers = expand_citation_numbers(text)
    return bool(numbers) and all(number in valid_refs for number in numbers)


def is_plain_citation_span(text: str, start: int, end: int, valid_refs: set[int]) -> bool:
    if start == 0 or not text[start - 1].isalnum():
        return False
    next_char = text[end : end + 1]
    if next_char and next_char.isalnum():
        return False
    if start > 0 and text[start - 1 : start] in {"/", "."}:
        return False
    token_match = re.search(r"[A-Za-z0-9./]+$", text[:start])
    token = token_match.group() if token_match else ""
    if any(char.isdigit() or char.isupper() for char in token) and not TOKEN_ENDING_WITH_BIO_NUMBER_RE.search(text[:start]):
        return False
    return is_valid_citation(text[start:end], valid_refs)


def plain_citation_spans(text: str, valid_refs: set[int]) -> list[tuple[int, int]]:
    spans: list[tuple[int, int]] = []
    cursor = 0
    while cursor < len(text):
        if not text[cursor].isdigit():
            cursor += 1
            continue
        best: tuple[int, int] | None = None
        for start in range(cursor, len(text)):
            if not text[start].isdigit():
                break
            match = NUMERIC_CITATION_RE.match(text, start)
            if not match:
                continue
            end = match.end()
            if not is_plain_citation_span(text, start, end, valid_refs):
                continue
            if start == cursor:
                best = (start, end)
            elif TOKEN_ENDING_WITH_BIO_NUMBER_RE.search(text[:start]):
                best = (start, end)
        if best is None:
            cursor += 1
            continue
        spans.append(best)
        cursor = best[1]
    return spans


def run_text_with_superscripts_masked(paragraph: ET.Element) -> str:
    pieces: list[str] = []
    for run in paragraph.findall(f"{W}r"):
        if is_superscript_run(run):
            pieces.append(" ")
        else:
            pieces.append(text_of(run))
    return "".join(pieces)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx")
    args = parser.parse_args()

    with zipfile.ZipFile(args.docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(f"{W}body")
    if body is None:
        raise RuntimeError("word/document.xml has no body")
    paragraphs = [p for p in body.findall(f"{W}p") if text_of(p).strip()]

    ref_index = len(paragraphs)
    valid_refs: set[int] = set()
    for index, paragraph in enumerate(paragraphs):
        text = text_of(paragraph).strip()
        match = REF_START_RE.match(text)
        if match:
            ref_index = min(ref_index, index)
            valid_refs.add(int(match.group(1)))
    for paragraph in paragraphs[ref_index + 1 :]:
        match = REF_START_RE.match(text_of(paragraph).strip())
        if match:
            valid_refs.add(int(match.group(1)))

    failures: list[tuple[int, str]] = []
    for index, paragraph in enumerate(paragraphs[:ref_index], start=1):
        masked = run_text_with_superscripts_masked(paragraph)
        for start, end in plain_citation_spans(masked, valid_refs):
            failures.append((index, masked[max(0, start - 20) : min(len(masked), end + 20)].strip()))

    if failures:
        print("FAIL: likely numeric citations remain as plain text.")
        for paragraph, text in failures[:50]:
            print(f"Paragraph {paragraph}: {text}")
        raise SystemExit(1)
    print("PASS: no likely plain-text numeric citations remain.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
