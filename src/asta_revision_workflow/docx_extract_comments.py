#!/usr/bin/env python3
"""Extract Word comments with their anchored paragraphs from a DOCX file."""

from __future__ import annotations

import argparse
import json
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"
NS = {"w": W_URI}


@dataclass(frozen=True)
class CommentAnchor:
    comment_id: str
    comment_text: str
    paragraph_index: int
    paragraph_text: str
    anchored_text: str


def text_of(elem: ET.Element) -> str:
    return "".join(t.text or "" for t in elem.findall(".//w:t", NS))


def comment_id(elem: ET.Element) -> str:
    return elem.attrib.get(f"{W}id", "")


def paragraph_comment_ids(paragraph: ET.Element) -> list[str]:
    ids: list[str] = []
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        for elem in paragraph.findall(f".//w:{tag}", NS):
            cid = comment_id(elem)
            if cid and cid not in ids:
                ids.append(cid)
    return ids


def anchored_text_for_comment(paragraph: ET.Element, cid: str) -> str:
    pieces: list[str] = []
    active = False
    for child in paragraph:
        if child.tag == f"{W}commentRangeStart" and comment_id(child) == cid:
            active = True
            continue
        if child.tag == f"{W}commentRangeEnd" and comment_id(child) == cid:
            active = False
            continue
        if active:
            pieces.append(text_of(child))
    anchored = "".join(pieces).strip()
    return anchored or text_of(paragraph).strip()


def extract_comments(docx: Path) -> list[CommentAnchor]:
    with zipfile.ZipFile(docx) as zf:
        names = set(zf.namelist())
        if "word/comments.xml" not in names:
            return []
        comments_root = ET.fromstring(zf.read("word/comments.xml"))
        document_root = ET.fromstring(zf.read("word/document.xml"))

    comments = {
        comment_id(elem): text_of(elem).strip()
        for elem in comments_root.findall("w:comment", NS)
        if comment_id(elem)
    }

    anchors: list[CommentAnchor] = []
    paragraphs = document_root.findall(".//w:body/w:p", NS)
    for index, paragraph in enumerate(paragraphs, start=1):
        paragraph_text = text_of(paragraph).strip()
        if not paragraph_text:
            continue
        for cid in paragraph_comment_ids(paragraph):
            anchors.append(
                CommentAnchor(
                    comment_id=cid,
                    comment_text=comments.get(cid, ""),
                    paragraph_index=index,
                    paragraph_text=paragraph_text,
                    anchored_text=anchored_text_for_comment(paragraph, cid),
                )
            )
    return anchors


def format_markdown(anchors: list[CommentAnchor]) -> str:
    chunks: list[str] = []
    for anchor in anchors:
        chunks.append(
            "\n".join(
                [
                    f"## Comment {anchor.comment_id}",
                    "",
                    f"Paragraph: {anchor.paragraph_index}",
                    "",
                    "Comment:",
                    anchor.comment_text,
                    "",
                    "Anchored text:",
                    anchor.anchored_text,
                    "",
                    "Paragraph text:",
                    anchor.paragraph_text,
                ]
            )
        )
    return "\n\n".join(chunks) + ("\n" if chunks else "")


def format_text(anchors: list[CommentAnchor]) -> str:
    chunks: list[str] = []
    for anchor in anchors:
        chunks.append(
            "\n".join(
                [
                    f"COMMENT {anchor.comment_id}",
                    f"PARAGRAPH: {anchor.paragraph_index}",
                    f"COMMENT TEXT: {anchor.comment_text}",
                    f"ANCHORED TEXT: {anchor.anchored_text}",
                    f"PARAGRAPH TEXT: {anchor.paragraph_text}",
                ]
            )
        )
    return "\n\n".join(chunks) + ("\n" if chunks else "")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx")
    parser.add_argument("-o", "--output", help="Write output to this file instead of stdout.")
    parser.add_argument(
        "--format",
        choices=["text", "markdown", "json"],
        default="text",
        help="Output format.",
    )
    args = parser.parse_args()

    anchors = extract_comments(Path(args.docx))
    if args.format == "json":
        output = json.dumps([asdict(anchor) for anchor in anchors], indent=2, ensure_ascii=False) + "\n"
    elif args.format == "markdown":
        output = format_markdown(anchors)
    else:
        output = format_text(anchors)

    if args.output:
        Path(args.output).write_text(output, encoding="utf-8")
    else:
        print(output, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
