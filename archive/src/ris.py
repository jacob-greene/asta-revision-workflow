from __future__ import annotations

import re
import sys
from pathlib import Path


def matching_brace(text: str, open_pos: int) -> int:
    depth = 0
    escaped = False
    for i in range(open_pos, len(text)):
        char = text[i]
        if escaped:
            escaped = False
            continue
        if char == "\\":
            escaped = True
            continue
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return i
    return -1


def split_top_level(text: str) -> list[str]:
    chunks: list[str] = []
    depth = 0
    quote = False
    start = 0
    for i, char in enumerate(text):
        if char == '"' and (i == 0 or text[i - 1] != "\\"):
            quote = not quote
        elif not quote:
            if char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
            elif char == "," and depth == 0:
                chunks.append(text[start:i].strip())
                start = i + 1
    tail = text[start:].strip()
    if tail:
        chunks.append(tail)
    return chunks


def clean_value(value: str) -> str:
    value = value.strip().rstrip(",").strip()
    if len(value) >= 2 and value[0] == "{" and value[-1] == "}":
        value = value[1:-1]
    elif len(value) >= 2 and value[0] == '"' and value[-1] == '"':
        value = value[1:-1]
    value = re.sub(r"\\&", "&", value)
    value = re.sub(r"\\['`^\"~=.]?\{?([A-Za-z])\}?", r"\1", value)
    value = re.sub(r"[{}]", "", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def parse_bibtex(text: str) -> list[tuple[str, str, dict[str, str]]]:
    entries: list[tuple[str, str, dict[str, str]]] = []
    pos = 0
    while True:
        start = text.find("@", pos)
        if start == -1:
            break
        type_end = text.find("{", start)
        if type_end == -1:
            break
        close = matching_brace(text, type_end)
        if close == -1:
            break

        entry_type = text[start + 1 : type_end].strip().lower()
        body = text[type_end + 1 : close]
        comma = body.find(",")
        if comma != -1:
            key = body[:comma].strip()
            fields: dict[str, str] = {}
            for chunk in split_top_level(body[comma + 1 :]):
                if "=" not in chunk:
                    continue
                name, value = chunk.split("=", 1)
                fields[name.strip().lower()] = clean_value(value)
            entries.append((entry_type, key, fields))
        pos = close + 1
    return entries


def ris_type(entry_type: str) -> str:
    return {
        "article": "JOUR",
        "inproceedings": "CONF",
        "proceedings": "CONF",
        "book": "BOOK",
        "phdthesis": "THES",
        "mastersthesis": "THES",
        "misc": "GEN",
    }.get(entry_type, "GEN")


def format_author(author: str) -> str:
    author = clean_value(author)
    if "," in author:
        return author
    parts = author.split()
    if len(parts) <= 1:
        return author
    return f"{parts[-1]}, {' '.join(parts[:-1])}"


def write_record(entry_type: str, key: str, fields: dict[str, str]) -> list[str]:
    lines = [f"TY  - {ris_type(entry_type)}"]
    if "title" in fields:
        lines.append(f"TI  - {fields['title']}")
    for author in re.split(r"\s+and\s+", fields.get("author", ""), flags=re.IGNORECASE):
        if author.strip():
            lines.append(f"AU  - {format_author(author)}")
    if "year" in fields:
        lines.append(f"PY  - {fields['year']}")
    if "journal" in fields:
        lines.append(f"JO  - {fields['journal']}")
    if "volume" in fields:
        lines.append(f"VL  - {fields['volume']}")
    if "number" in fields:
        lines.append(f"IS  - {fields['number']}")
    if "pages" in fields:
        pages = fields["pages"].replace("--", "-")
        if "-" in pages:
            start, end = pages.split("-", 1)
            lines.append(f"SP  - {start}")
            lines.append(f"EP  - {end}")
        else:
            lines.append(f"SP  - {pages}")
    if "doi" in fields:
        lines.append(f"DO  - {fields['doi']}")
    if "url" in fields:
        lines.append(f"UR  - {fields['url']}")
    lines.append(f"ID  - {key}")
    lines.append("ER  -")
    return lines


def bibtex_to_ris(text: str) -> str:
    records: list[str] = []
    for entry_type, key, fields in parse_bibtex(text):
        records.extend(write_record(entry_type, key, fields))
        records.append("")
    return "\n".join(records)


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: bib-to-ris input.bib output.ris", file=sys.stderr)
        return 2

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    entries = parse_bibtex(input_path.read_text(encoding="utf-8"))
    records: list[str] = []
    for entry_type, key, fields in entries:
        records.extend(write_record(entry_type, key, fields))
        records.append("")
    output_path.write_text("\n".join(records), encoding="utf-8")
    print(f"Wrote {len(entries)} records to {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
