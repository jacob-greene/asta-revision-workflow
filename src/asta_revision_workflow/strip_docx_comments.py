#!/usr/bin/env python3
"""Copy a DOCX while removing Word comments and their inline anchors."""

from __future__ import annotations

import argparse
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_URI = "http://schemas.openxmlformats.org/package/2006/relationships"
W = f"{{{W_URI}}}"
REL = f"{{{REL_URI}}}"

ET.register_namespace("w", W_URI)

COMMENT_PARTS = {
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsExtensible.xml",
    "word/commentsIds.xml",
}


def clear_comment_markup(elem: ET.Element) -> None:
    for parent in elem.iter():
        for child in list(parent):
            tag = child.tag.split("}")[-1]
            if tag in {"commentRangeStart", "commentRangeEnd"}:
                parent.remove(child)
            elif tag == "r":
                refs = [x for x in list(child) if x.tag.split("}")[-1] == "commentReference"]
                if refs and len(list(child)) == len(refs):
                    parent.remove(child)
                else:
                    for ref in refs:
                        child.remove(ref)


def remove_comment_relationships(rels_path: Path) -> None:
    if not rels_path.exists():
        return
    tree = ET.parse(rels_path)
    root = tree.getroot()
    for rel in list(root):
        rel_type = rel.attrib.get("Type", "")
        target = rel.attrib.get("Target", "")
        if rel_type.endswith("/comments") or target in {
            "comments.xml",
            "commentsExtended.xml",
            "commentsExtensible.xml",
            "commentsIds.xml",
        }:
            root.remove(rel)
    tree.write(rels_path, encoding="UTF-8", xml_declaration=True)


def remove_content_type_overrides(content_types_path: Path) -> None:
    if not content_types_path.exists():
        return
    tree = ET.parse(content_types_path)
    root = tree.getroot()
    for child in list(root):
        part_name = child.attrib.get("PartName", "").lstrip("/")
        if part_name in COMMENT_PARTS:
            root.remove(child)
    tree.write(content_types_path, encoding="UTF-8", xml_declaration=True)


def strip_comments(source: Path, output: Path) -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(source) as zf:
            zf.extractall(tmp)

        document = tmp / "word" / "document.xml"
        tree = ET.parse(document)
        clear_comment_markup(tree.getroot())
        tree.write(document, encoding="UTF-8", xml_declaration=True)

        for part in COMMENT_PARTS:
            path = tmp / part
            if path.exists():
                path.unlink()
        remove_comment_relationships(tmp / "word" / "_rels" / "document.xml.rels")
        remove_content_type_overrides(tmp / "[Content_Types].xml")

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
    args = parser.parse_args()
    strip_comments(Path(args.source), Path(args.output))
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
