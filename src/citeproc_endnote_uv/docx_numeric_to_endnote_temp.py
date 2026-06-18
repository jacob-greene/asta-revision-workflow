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

from citeproc_endnote_uv.strip_docx_comments import (
    COMMENT_PARTS,
    remove_comment_relationships,
    remove_content_type_overrides,
)

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"
ET.register_namespace("w", W_URI)

CITE_RE = re.compile(r"^\s*\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*\s*$")
REF_START_RE = re.compile(r"(?<![A-Za-z0-9])(\d{1,3})\.\s*(?=[A-Z])")
TOKEN_ENDING_WITH_BIO_NUMBER_RE = re.compile(
    r"(?:PRC\d+|H\d(?:\.\d)?K27(?:me\d)?|H\d(?:\.\d)?K27M|H\d(?:\.\d)?)\.$"
)


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


def reference_blocks(paragraphs: list[ET.Element]) -> list[tuple[int, int, int, int]]:
    blocks: list[tuple[int, int, int, int]] = []
    for index, paragraph in enumerate(paragraphs):
        if not re.match(r"^\s*1\.\s*[A-Z]", text_of(paragraph).strip()):
            continue
        end = index
        last_number = 0
        max_number = 0
        count = 0
        for candidate in paragraphs[index:]:
            text = re.sub(r"\s+", " ", text_of(candidate).strip())
            match = REF_START_RE.match(text)
            if count and not match:
                break
            if match:
                number = int(match.group(1))
                if count and number <= last_number:
                    break
                last_number = number
                max_number = max(max_number, number)
                count += 1
                end += 1
        if count:
            blocks.append((index, end, max_number, count))
    return blocks


def selected_reference_block(paragraphs: list[ET.Element]) -> tuple[int, int]:
    blocks = reference_blocks(paragraphs)
    if not blocks:
        raise RuntimeError("Could not locate numbered reference list in DOCX.")
    start, end, _, _ = max(blocks, key=lambda block: (block[2], block[3]))
    return start, end


def parse_ris_references(path: Path) -> dict[int, Reference]:
    mapping: dict[int, Reference] = {}
    current: dict[str, str] = {}
    authors: list[str] = []
    index = 1
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith("ER  -"):
            if authors and current.get("PY"):
                first = authors[0].split(",", 1)[0].strip()
                mapping[index] = Reference(
                    author=first,
                    year=current.get("PY", ""),
                    title=current.get("TI", ""),
                )
                index += 1
            current = {}
            authors = []
            continue
        match = re.match(r"([A-Z0-9]{2})  - (.*)", line)
        if not match:
            continue
        key, value = match.group(1), re.sub(r"\s+", " ", match.group(2)).strip()
        if key == "AU":
            authors.append(value)
        else:
            current[key] = value
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
    if len(after_year) >= 2 and after_year[1].strip():
        title = after_year[1].split(".", 1)[0].strip()
        return re.sub(r"\s+", " ", title)

    if len(after_year) < 2 or not after_year[1].strip():
        end_year = re.search(r"\((\d{4})\)\.?$", entry)
        if not end_year:
            return ""
        before_year = entry[: end_year.start()].strip()
        if " et al. " in before_year:
            _, rest = before_year.split(" et al. ", 1)
        elif ". " in before_year:
            _, rest = before_year.split(". ", 1)
        else:
            return ""
        title = rest.split(". ", 1)[0].strip().rstrip(".")
        return re.sub(r"\s+", " ", title)


def normalized_title(title: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", title.lower()).strip()


TITLE_TRAILING_STOPWORDS = {
    "a",
    "an",
    "and",
    "as",
    "by",
    "for",
    "from",
    "in",
    "of",
    "on",
    "or",
    "the",
    "to",
    "with",
}


def short_title_key(title: str, max_words: int = 10) -> str:
    words = re.findall(r"[A-Za-z0-9]+", title)
    if not words:
        return ""
    prefix = words[:max_words]
    while len(prefix) > 1 and prefix[-1].lower() in TITLE_TRAILING_STOPWORDS:
        prefix.pop()
    return " ".join(prefix)


def unique_short_title_key(title: str, peer_titles: list[str], min_words: int = 10) -> str:
    words = re.findall(r"[A-Za-z0-9]+", title)
    if not words:
        return ""
    start = min(max(1, min_words), len(words))
    normalized_peers = [normalized_title(peer) for peer in peer_titles if peer]
    for count in range(start, len(words) + 1):
        prefix = short_title_key(title, max_words=count)
        normalized_prefix = normalized_title(prefix)
        exact_matches = [peer for peer in normalized_peers if peer == normalized_prefix]
        if len(exact_matches) == 1:
            return prefix
        prefix_matches = [peer for peer in normalized_peers if peer.startswith(normalized_prefix)]
        if len(set(prefix_matches)) <= 1:
            return prefix
    return title


def temporary_citation_text(
    reference: Reference,
    ambiguous: bool,
    disambiguate_with_title: bool = True,
    short_title_disambiguation: bool = False,
    title_prefix_words: int = 10,
    title_override: str | None = None,
) -> str:
    if disambiguate_with_title and ambiguous and reference.title:
        title = (
            title_override or short_title_key(reference.title, max_words=title_prefix_words)
            if short_title_disambiguation
            else reference.title
        )
        return f"{reference.author}, {reference.year}, {title}"
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


def make_temp_citation(
    text: str,
    references: dict[int, Reference],
    ambiguous_keys: set[str],
    disambiguate_with_title: bool = True,
    short_title_disambiguation: bool = False,
    title_prefix_words: int = 10,
    title_overrides: dict[int, str] | None = None,
) -> str:
    parts: list[str] = []
    for number in expand_citation_numbers(text):
        reference = references.get(number)
        if reference is None:
            parts.append(f"REF{number}")
        else:
            parts.append(
                temporary_citation_text(
                    reference,
                    reference.author_year in ambiguous_keys,
                    disambiguate_with_title=disambiguate_with_title,
                    short_title_disambiguation=short_title_disambiguation,
                    title_prefix_words=title_prefix_words,
                    title_override=(title_overrides or {}).get(number),
                )
            )
    return "{" + "; ".join(parts) + "}"


def is_biological_decimal_label(preceding: str, run_text: str) -> bool:
    return (
        bool(re.fullmatch(r"\s*\d+\s*", run_text))
        and bool(TOKEN_ENDING_WITH_BIO_NUMBER_RE.search(preceding.rstrip()))
    )


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


def run_contains_endnote_instr(run: ET.Element) -> bool:
    return any("ADDIN EN." in (instr.text or "") for instr in run.findall(f".//{W}instrText"))


def run_field_char_type(run: ET.Element) -> str | None:
    fld_char = run.find(f"{W}fldChar")
    if fld_char is None:
        return None
    return fld_char.attrib.get(f"{W}fldCharType")


def flatten_endnote_fields(root: ET.Element) -> int:
    """Remove old EndNote field wrappers while preserving their displayed results."""

    flattened = 0
    for paragraph in root.findall(f".//{W}p"):
        children = list(paragraph)
        index = 0
        while index < len(children):
            child = children[index]
            if child.tag != f"{W}r" or run_field_char_type(child) != "begin":
                index += 1
                continue

            depth = 0
            end_index = None
            separate_index = None
            has_endnote_instr = False
            for probe in range(index, len(children)):
                probe_child = children[probe]
                if probe_child.tag != f"{W}r":
                    continue
                char_type = run_field_char_type(probe_child)
                if char_type == "begin":
                    depth += 1
                elif char_type == "separate" and depth == 1 and separate_index is None:
                    separate_index = probe
                elif char_type == "end":
                    depth -= 1
                    if depth == 0:
                        end_index = probe
                        break
                if run_contains_endnote_instr(probe_child):
                    has_endnote_instr = True

            if not has_endnote_instr or end_index is None or separate_index is None:
                index += 1
                continue

            remove_indices = set(range(index, separate_index + 1))
            remove_indices.add(end_index)
            for remove_index in sorted(remove_indices, reverse=True):
                paragraph.remove(children[remove_index])
            flattened += 1
            children = list(paragraph)
            index = index
    flattened += remove_endnote_field_control_runs(root)
    return flattened


def remove_endnote_field_control_runs(root: ET.Element) -> int:
    """Remove EndNote field-code runs that can span multiple paragraphs.

    EndNote bibliographies may be represented as one field beginning in the
    first reference paragraph and ending in the last reference paragraph. The
    displayed reference text is already plain text in separate runs, so removing
    field-control and instruction runs preserves the visible bibliography while
    removing stale EndNote state.
    """

    removed = 0
    in_endnote_field = False
    pending_begin: tuple[ET.Element, ET.Element] | None = None
    for paragraph in root.findall(f".//{W}p"):
        for run in list(paragraph):
            if run.tag != f"{W}r":
                continue
            char_type = run_field_char_type(run)
            instr = "".join(instr.text or "" for instr in run.findall(f".//{W}instrText"))
            if char_type == "begin":
                pending_begin = (paragraph, run)
                continue
            if "ADDIN EN." in instr:
                if pending_begin is not None:
                    begin_parent, begin_run = pending_begin
                    if begin_run in list(begin_parent):
                        begin_parent.remove(begin_run)
                        removed += 1
                pending_begin = None
                in_endnote_field = True
                paragraph.remove(run)
                removed += 1
                continue
            if in_endnote_field and char_type in {"separate", "end"}:
                paragraph.remove(run)
                removed += 1
                if char_type == "end":
                    in_endnote_field = False
                continue
            if pending_begin is not None and run is not pending_begin[1]:
                pending_begin = None
    return removed


def remove_endnote_docvars(settings_path: Path) -> int:
    if not settings_path.exists():
        return 0
    tree = ET.parse(settings_path)
    root = tree.getroot()
    removed = 0
    for doc_vars in root.findall(f"{W}docVars"):
        for doc_var in list(doc_vars):
            if doc_var.attrib.get(f"{W}name", "").startswith("EN."):
                doc_vars.remove(doc_var)
                removed += 1
        if len(list(doc_vars)) == 0:
            root.remove(doc_vars)
    if removed:
        tree.write(settings_path, encoding="UTF-8", xml_declaration=True)
    return removed


def remove_endnote_styles(styles_path: Path) -> int:
    if not styles_path.exists():
        return 0
    tree = ET.parse(styles_path)
    root = tree.getroot()
    removed = 0
    for style in list(root.findall(f"{W}style")):
        style_id = style.attrib.get(f"{W}styleId", "")
        name = style.find(f"{W}name")
        style_name = name.attrib.get(f"{W}val", "") if name is not None else ""
        if "EndNote" in style_id or "EndNote" in style_name:
            root.remove(style)
            removed += 1
    if removed:
        tree.write(styles_path, encoding="UTF-8", xml_declaration=True)
    return removed


def remove_comment_parts(package_dir: Path) -> None:
    for part in COMMENT_PARTS:
        path = package_dir / part
        if path.exists():
            path.unlink()
    remove_comment_relationships(package_dir / "word" / "_rels" / "document.xml.rels")
    remove_content_type_overrides(package_dir / "[Content_Types].xml")


def convert(
    source: Path,
    output: Path,
    keep_references: bool = False,
    disambiguate_with_title: bool = True,
    short_title_disambiguation: bool = False,
    title_prefix_words: int = 10,
    ris: Path | None = None,
) -> None:
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
        ref_index, ref_end = selected_reference_block(paragraphs)
        first_ref_index = min(start for start, _, _, _ in reference_blocks(paragraphs))

        reference_text = " ".join(text_of(p) for p in paragraphs[ref_index:ref_end])
        references = parse_ris_references(ris) if ris is not None else parse_references(reference_text)
        if not references:
            raise RuntimeError("Could not parse numbered references.")
        author_year_counts: dict[str, int] = {}
        for reference in references.values():
            author_year_counts[reference.author_year] = author_year_counts.get(reference.author_year, 0) + 1
        ambiguous_keys = {key for key, count in author_year_counts.items() if count > 1}
        titles_by_author_year: dict[str, list[str]] = {}
        for reference in references.values():
            titles_by_author_year.setdefault(reference.author_year, []).append(reference.title)
        title_overrides = {
            number: unique_short_title_key(reference.title, titles_by_author_year.get(reference.author_year, []), title_prefix_words)
            for number, reference in references.items()
            if reference.author_year in ambiguous_keys and reference.title
        }

        for paragraph in paragraphs[:first_ref_index]:
            preceding = ""
            for run in paragraph.findall(f"{W}r"):
                run_text = text_of(run)
                if is_superscript_run(run) and CITE_RE.match(run_text):
                    if is_biological_decimal_label(preceding, run_text):
                        set_run_text(run, run_text.strip())
                        preceding += run_text.strip()
                        continue
                    citation = make_temp_citation(
                        run_text,
                        references,
                        ambiguous_keys,
                        disambiguate_with_title=disambiguate_with_title,
                        short_title_disambiguation=short_title_disambiguation,
                        title_prefix_words=title_prefix_words,
                        title_overrides=title_overrides,
                    )
                    if preceding and not preceding[-1].isspace():
                        citation = " " + citation
                    set_run_text(run, citation)
                    preceding += citation
                else:
                    preceding += run_text

        flatten_endnote_fields(root)

        if not keep_references:
            for paragraph in paragraphs[first_ref_index:]:
                parent_children = list(body)
                if paragraph in parent_children:
                    body.remove(paragraph)
        else:
            selected = set(paragraphs[ref_index:ref_end])
            for start, end, _, _ in reference_blocks(paragraphs):
                for paragraph in paragraphs[start:end]:
                    if paragraph in selected:
                        continue
                    parent_children = list(body)
                    if paragraph in parent_children:
                        body.remove(paragraph)

        tree.write(document, encoding="UTF-8", xml_declaration=True)
        remove_endnote_docvars(tmp / "word" / "settings.xml")
        remove_endnote_styles(tmp / "word" / "styles.xml")
        remove_comment_parts(tmp)
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
    parser.add_argument("--ris", help="Use this RIS, in numbered-reference order, as citation metadata.")
    parser.add_argument("--keep-references", action="store_true")
    parser.add_argument(
        "--no-title-disambiguation",
        action="store_true",
        help="Use only {Author, Year} temporary citations, even for ambiguous author-year pairs. This is less precise but more robust for EndNote parsing.",
    )
    parser.add_argument(
        "--short-title-disambiguation",
        action="store_true",
        help="For ambiguous author-year pairs, include a readable title prefix instead of the full title.",
    )
    parser.add_argument(
        "--title-prefix-words",
        type=int,
        default=10,
        help="Maximum title words to include with --short-title-disambiguation. Trailing connector words are removed.",
    )
    args = parser.parse_args()
    convert(
        Path(args.source),
        Path(args.output),
        keep_references=args.keep_references,
        disambiguate_with_title=not args.no_title_disambiguation,
        short_title_disambiguation=args.short_title_disambiguation,
        title_prefix_words=args.title_prefix_words,
        ris=Path(args.ris) if args.ris else None,
    )
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
