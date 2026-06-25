from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Any


Inline = dict[str, Any]
BibEntry = dict[str, str]
INLINE_REPLACEMENT = "__citeproc_endnote_inline_replacement__"


def main() -> int:
    document = json.load(sys.stdin)
    bibliography = load_bibliography(find_bibliography(document.get("meta", {})))
    transformed = walk(document, bibliography)
    json.dump(transformed, sys.stdout, ensure_ascii=False)
    return 0


def find_bibliography(meta: dict[str, Any]) -> Path | None:
    explicit = os.environ.get("CITEPROC_ENDNOTE_BIBLIOGRAPHY")
    if explicit:
        return Path(explicit)

    for path in bibliography_paths_from_meta(meta):
        if path.exists():
            return path

    for name in ("references.bib", "latex/references.bib"):
        path = Path(name)
        if path.exists():
            return path

    return None


def bibliography_paths_from_meta(meta: dict[str, Any]) -> list[Path]:
    raw = meta.get("bibliography")
    if not raw:
        return []

    values: list[Any]
    if isinstance(raw, dict) and raw.get("t") == "MetaList":
        values = raw.get("c", [])
    else:
        values = [raw]

    paths: list[Path] = []
    for value in values:
        text = meta_value_to_text(value)
        if text:
            paths.append(Path(text))
    return paths


def meta_value_to_text(value: Any) -> str:
    if isinstance(value, dict):
        tag = value.get("t")
        content = value.get("c")
        if tag in {"MetaString", "MetaInlines", "MetaBlocks"}:
            return stringify(content)
    if isinstance(value, str):
        return value
    return ""


def load_bibliography(path: Path | None) -> dict[str, BibEntry]:
    if path is None or not path.exists():
        return {}
    return parse_bibtex(path.read_text(encoding="utf-8"))


def parse_bibtex(text: str) -> dict[str, BibEntry]:
    entries: dict[str, BibEntry] = {}
    pos = 0
    while True:
        start = text.find("@", pos)
        if start == -1:
            break
        open_pos = text.find("{", start)
        if open_pos == -1:
            break
        close_pos = matching_brace(text, open_pos)
        if close_pos == -1:
            break

        body = text[open_pos + 1 : close_pos]
        key, fields = parse_entry_body(body)
        if key:
            entries[key] = fields
        pos = close_pos + 1
    return entries


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


def parse_entry_body(body: str) -> tuple[str, BibEntry]:
    comma = find_top_level_comma(body)
    if comma == -1:
        return "", {}
    key = body[:comma].strip()
    fields: BibEntry = {}
    for chunk in split_top_level(body[comma + 1 :]):
        if "=" not in chunk:
            continue
        name, value = chunk.split("=", 1)
        fields[name.strip().lower()] = clean_bibtex_value(value)
    return key, fields


def find_top_level_comma(text: str) -> int:
    depth = 0
    quote = False
    for i, char in enumerate(text):
        if char == '"' and (i == 0 or text[i - 1] != "\\"):
            quote = not quote
        elif not quote:
            if char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
            elif char == "," and depth == 0:
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


def clean_bibtex_value(value: str) -> str:
    value = value.strip().rstrip(",").strip()
    if len(value) >= 2 and value[0] == "{" and value[-1] == "}":
        value = value[1:-1]
    elif len(value) >= 2 and value[0] == '"' and value[-1] == '"':
        value = value[1:-1]
    value = re.sub(r"[{}]", "", value)
    value = re.sub(r"\\['`^\"~=.]?\{?([A-Za-z])\}?", r"\1", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def walk(node: Any, bibliography: dict[str, BibEntry]) -> Any:
    if isinstance(node, dict):
        if node.get("t") == "Cite":
            return {INLINE_REPLACEMENT: citation_to_inlines(node, bibliography)}
        return {key: walk(value, bibliography) for key, value in node.items()}
    if isinstance(node, list):
        out: list[Any] = []
        for value in node:
            new_value = walk(value, bibliography)
            if isinstance(new_value, dict) and INLINE_REPLACEMENT in new_value:
                out.extend(new_value[INLINE_REPLACEMENT])
            else:
                out.append(new_value)
        return out
    return node


def citation_to_inlines(node: Inline, bibliography: dict[str, BibEntry]) -> list[Inline]:
    citations = node.get("c", [[], []])[0]
    before: list[Inline] = []
    after: list[Inline] = []
    cite_parts: list[str] = []

    for citation in citations:
        before.extend(citation.get("citationPrefix", []))
        citation_id = citation.get("citationId", "")
        cite_parts.append(endnote_citation(citation_id, bibliography.get(citation_id, {})))
        after.extend(citation.get("citationSuffix", []))

    output: list[Inline] = []
    if before:
        output.extend(before)
        output.append({"t": "Space"})
    output.append({"t": "Str", "c": "{" + "; ".join(cite_parts) + "}"})
    if after:
        output.append({"t": "Space"})
        output.extend(after)
    return output


def endnote_citation(citation_id: str, entry: BibEntry) -> str:
    author = first_author(entry.get("author", ""))
    year = year_from_entry(entry)
    record_number = entry.get("record-number") or entry.get("record_number") or entry.get("endnote")

    if author and year and record_number:
        return f"{author}, {year} #{record_number}"
    if author and year:
        return f"{author}, {year}"
    return citation_id


def first_author(author_field: str) -> str:
    if not author_field:
        return ""
    first = re.split(r"\s+and\s+", author_field, maxsplit=1, flags=re.IGNORECASE)[0].strip()
    if "," in first:
        return clean_author(first.split(",", 1)[0])
    parts = clean_author(first).split()
    return parts[-1] if parts else ""


def clean_author(author: str) -> str:
    author = re.sub(r"[{}]", "", author)
    author = re.sub(r"\s+", " ", author)
    return author.strip()


def year_from_entry(entry: BibEntry) -> str:
    for field in ("year", "date"):
        match = re.search(r"\d{4}", entry.get(field, ""))
        if match:
            return match.group(0)
    return ""


def stringify(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, list):
        return "".join(stringify(item) for item in value)
    if isinstance(value, dict):
        tag = value.get("t")
        content = value.get("c")
        if tag == "Str":
            return str(content)
        if tag == "Space":
            return " "
        if content is not None:
            return stringify(content)
    return ""


if __name__ == "__main__":
    raise SystemExit(main())
