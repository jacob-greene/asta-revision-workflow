#!/usr/bin/env python3
"""Launch and finalize a revision pass whose only content input is one Word DOCX.

The launcher intentionally rejects bibliography, RIS, archive, Markdown, TeX,
and cached evidence inputs. It creates a run directory from the source DOCX,
extracts comments/reference metadata from that DOCX, and records a manifest.
The finalize step accepts only files inside the run directory produced from the
same manifest.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"
NS = {"w": W_URI}
SCRIPT_DIR = Path(__file__).resolve().parent
COMMENT_PARTS = {
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsExtensible.xml",
    "word/commentsIds.xml",
}
REF_START_RE = re.compile(r"^\s*(\d{1,3})\.\s*(?=[A-Z])")
CITATION_CLUSTER_RE = re.compile(r"(?<![A-Za-z0-9.])\d{1,3}(?:\s*-\s*\d{1,3})?(?:\s*,\s*\d{1,3}(?:\s*-\s*\d{1,3})?)*(?![A-Za-z0-9])")
REVIEWER_TASKS = {
    "scientific_rigor_reviewer.md": (
        "Review only the launcher-scoped revised paragraphs. Be skeptical of new knowledge claims, "
        "unsupported causal language, overgeneralization across lineages, and claims that simply restate earlier text. "
        "Report paragraph-specific findings with severity and required fixes."
    ),
    "citation_ris_reviewer.md": (
        "Review citation support and bibliography integrity. Check that each modified statement is supported by "
        "same/adjacent citations, that citation numbers map to the intended claims, and that the paired RIS has "
        "complete author metadata without et-al placeholders or missing AU fields."
    ),
    "style_reviewer.md": (
        "Review tone and paragraph logic against the scientific-writing skills. Check topic sentences, concise "
        "claim-evidence-implication flow, non-redundancy with nearby paragraphs, and consistency with the review style."
    ),
}


@dataclass(frozen=True)
class CommentAnchor:
    comment_id: str
    comment_text: str
    paragraph_index: int
    paragraph_text: str
    anchored_text: str


def run(command: list[str]) -> None:
    print("+ " + " ".join(command))
    subprocess.run(command, check=True)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


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


def docx_roots(docx: Path) -> tuple[ET.Element, ET.Element | None]:
    with zipfile.ZipFile(docx) as zf:
        names = set(zf.namelist())
        document_root = ET.fromstring(zf.read("word/document.xml"))
        comments_root = ET.fromstring(zf.read("word/comments.xml")) if "word/comments.xml" in names else None
    return document_root, comments_root


def content_paragraphs(document_root: ET.Element) -> list[str]:
    paragraphs = []
    for paragraph in document_root.findall(".//w:body/w:p", NS):
        text = text_of(paragraph).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def extract_comments(docx: Path) -> list[CommentAnchor]:
    document_root, comments_root = docx_roots(docx)
    if comments_root is None:
        return []
    comments = {
        comment_id(elem): text_of(elem).strip()
        for elem in comments_root.findall("w:comment", NS)
        if comment_id(elem)
    }
    anchors: list[CommentAnchor] = []
    for index, paragraph in enumerate(document_root.findall(".//w:body/w:p", NS), start=1):
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


def reference_numbers(paragraphs: list[str]) -> list[int]:
    numbers = []
    for paragraph in paragraphs:
        match = REF_START_RE.match(paragraph)
        if match:
            numbers.append(int(match.group(1)))
    return numbers


def without_numeric_citation_clusters(text: str) -> str:
    return re.sub(r"\s+", " ", CITATION_CLUSTER_RE.sub("", text)).strip()


def is_citation_only_change(source_text: str, raw_text: str) -> bool:
    return without_numeric_citation_clusters(source_text) == without_numeric_citation_clusters(raw_text)


def reference_region_start(paragraphs: list[str]) -> int | None:
    for index, paragraph in enumerate(paragraphs, start=1):
        if REF_START_RE.match(paragraph):
            return index
    return None


def allowed_reference_integrity_changes(source_docx: Path, raw_docx: Path) -> list[int]:
    """Allow narrow bibliography repairs without opening prose outside comments."""

    source_paragraphs = content_paragraphs(docx_roots(source_docx)[0])
    raw_paragraphs = content_paragraphs(docx_roots(raw_docx)[0])
    source_start = reference_region_start(source_paragraphs)
    raw_start = reference_region_start(raw_paragraphs)
    if source_start is None or raw_start is None or source_start != raw_start:
        return []

    allowed: list[int] = []
    for index, (source_text, raw_text) in enumerate(zip(source_paragraphs, raw_paragraphs), start=1):
        if index < source_start or source_text == raw_text:
            continue
        if REF_START_RE.match(source_text) and REF_START_RE.match(raw_text):
            allowed.append(index)
    return allowed


def reference_duplicate_surplus(paragraphs: list[str]) -> int:
    from citeproc_endnote_uv import docx_reference_list_to_ris as word_refs

    seen: set[tuple[str, str]] = set()
    surplus = 0
    for paragraph in paragraphs:
        match = REF_START_RE.match(paragraph)
        if not match:
            continue
        ref = word_refs.parse_reference(int(match.group(1)), paragraph[match.end() :])
        if ref is None:
            continue
        year = ref.year
        title = ref.title
        key = (re.sub(r"[^a-z0-9]+", " ", title.lower()).strip(), year)
        if key in seen and key[0]:
            surplus += 1
        else:
            seen.add(key)
    return surplus


def write_comments_markdown(comments: list[CommentAnchor], output: Path) -> None:
    chunks = []
    for anchor in comments:
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
    output.write_text("\n\n".join(chunks) + ("\n" if chunks else ""), encoding="utf-8")


def write_reviewer_tasks(run_dir: Path, stem: str, manifest: dict) -> dict[str, list[str] | str]:
    tasks_dir = run_dir / "reviewer_tasks"
    reports_dir = run_dir / "reviewer_reports"
    tasks_dir.mkdir(exist_ok=True)
    reports_dir.mkdir(exist_ok=True)
    task_files: list[str] = []
    report_files: list[str] = []
    for filename, instruction in REVIEWER_TASKS.items():
        report_name = filename.replace("_reviewer.md", "_report.md")
        task_path = tasks_dir / filename
        report_path = reports_dir / report_name
        task_path.write_text(
            "\n".join(
                [
                    f"# {filename.removesuffix('.md').replace('_', ' ').title()}",
                    "",
                    instruction,
                    "",
                    "## Run Inputs",
                    f"- Source DOCX: `{manifest['source_docx']}`",
                    f"- Comments: `{manifest['generated_artifacts']['comments_markdown']}`",
                    f"- Allowed paragraphs: `{', '.join(str(i) for i in manifest['allowed_paragraphs'])}`",
                    "",
                    "## Run Outputs To Review",
                    f"- Raw revised DOCX: `{manifest['generated_artifacts']['raw_docx']}`",
                    f"- Final DOCX: `{manifest['generated_artifacts']['final_docx']}`",
                    f"- RIS: `{manifest['generated_artifacts']['ris']}`",
                    "",
                    "Write the report to:",
                    f"`{report_path.relative_to(run_dir)}`",
                    "",
                ]
            ),
            encoding="utf-8",
        )
        task_files.append(str(task_path.relative_to(run_dir)))
        report_files.append(str(report_path.relative_to(run_dir)))
    return {"tasks_dir": str(tasks_dir.relative_to(run_dir)), "task_files": task_files, "required_reports": report_files}


def safe_stem(path: Path) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", path.stem).strip("._") or "word_doc_revision"


def ensure_docx_only(path: Path) -> None:
    if path.suffix.lower() != ".docx":
        raise SystemExit(f"Only a .docx source is allowed: {path}")
    if not path.exists():
        raise SystemExit(f"Missing source DOCX: {path}")
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        if "word/document.xml" not in names:
            raise SystemExit(f"Not a valid Word DOCX package: {path}")


def start(args: argparse.Namespace) -> int:
    source = Path(args.source_docx).resolve()
    ensure_docx_only(source)

    stem = args.output_stem or safe_stem(source)
    run_dir = Path(args.run_dir or f"{stem}_word_doc_only_run").resolve()
    run_dir.mkdir(parents=True, exist_ok=False)

    source_copy = run_dir / source.name
    shutil.copy2(source, source_copy)
    document_root, _ = docx_roots(source_copy)
    paragraphs = content_paragraphs(document_root)
    comments = extract_comments(source_copy)
    allowed_paragraphs = sorted({anchor.paragraph_index for anchor in comments})
    refs = reference_numbers(paragraphs)

    comments_md = run_dir / f"{stem}.comments.md"
    comments_json = run_dir / f"{stem}.comments.json"
    manifest_path = run_dir / "manifest.json"
    raw_docx = run_dir / f"{stem}.raw.docx"
    final_docx = run_dir / f"{stem}.docx"
    ris = run_dir / f"{stem}.ris"

    write_comments_markdown(comments, comments_md)
    comments_json.write_text(
        json.dumps([asdict(comment) for comment in comments], indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    manifest = {
        "content_policy": "single-word-doc-only",
        "source_docx": source_copy.name,
        "source_sha256": sha256(source_copy),
        "source_paragraph_count": len(paragraphs),
        "comment_count": len(comments),
        "allowed_paragraphs": allowed_paragraphs,
        "reference_count": len(refs),
        "reference_numbers": refs,
        "forbidden_inputs": [
            "archive RIS",
            "external RIS",
            "BibTeX",
            "Markdown draft",
            "TeX draft",
            "prior response file",
            "cached Asta evidence",
            "hard-coded replacement paragraphs from older passes",
        ],
        "generated_artifacts": {
            "comments_markdown": comments_md.name,
            "comments_json": comments_json.name,
            "raw_docx": raw_docx.name,
            "final_docx": final_docx.name,
            "ris": ris.name,
        },
    }
    manifest["reviewers"] = write_reviewer_tasks(run_dir, stem, manifest)
    manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")

    print(f"Wrote run directory: {run_dir}")
    print(f"Only content source: {source_copy}")
    print(f"Allowed paragraph scope from comments: {allowed_paragraphs or 'none'}")
    print("Next: create the raw revised DOCX inside this run directory, then run finalize.")
    print(f"Finalize command: {Path(sys.argv[0]).name} finalize {manifest_path}")
    print("Required spawned reviewer task files:")
    for task in manifest["reviewers"]["task_files"]:
        print(f"  - {run_dir / task}")
    return 0


def stale_marker_counts(docx: Path) -> dict[str, int]:
    with zipfile.ZipFile(docx) as zf:
        data = "\n".join(zf.namelist()) + "\n"
        for name in zf.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                data += zf.read(name).decode("utf-8", "ignore")
    tokens = ["ADDIN EN.CITE", "EN.CWYW", "commentRangeStart", "commentRangeEnd", "commentReference"]
    counts = {token: data.count(token) for token in tokens}
    counts["comment_parts"] = sum(1 for part in COMMENT_PARTS if part in data)
    return counts


def paragraph_differences(source_docx: Path, raw_docx: Path) -> tuple[list[int], int, int]:
    source_paragraphs = content_paragraphs(docx_roots(source_docx)[0])
    raw_paragraphs = content_paragraphs(docx_roots(raw_docx)[0])
    changed = [
        index
        for index, (source_text, raw_text) in enumerate(zip(source_paragraphs, raw_paragraphs), start=1)
        if source_text != raw_text
    ]
    if len(raw_paragraphs) > len(source_paragraphs):
        changed.extend(range(len(source_paragraphs) + 1, len(raw_paragraphs) + 1))
    return changed, len(source_paragraphs), len(raw_paragraphs)


def ensure_inside_run_dir(path: Path, run_dir: Path, label: str) -> Path:
    resolved = path.resolve()
    try:
        resolved.relative_to(run_dir)
    except ValueError as exc:
        raise SystemExit(f"{label} must be inside the run directory: {resolved}") from exc
    return resolved


def finalize(args: argparse.Namespace) -> int:
    manifest_path = Path(args.manifest).resolve()
    run_dir = manifest_path.parent
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("content_policy") != "single-word-doc-only":
        raise SystemExit("Manifest is not a single-word-doc-only revision manifest.")

    source_docx = run_dir / manifest["source_docx"]
    if sha256(source_docx) != manifest["source_sha256"]:
        raise SystemExit("Source DOCX hash no longer matches the launch manifest.")

    artifacts = manifest["generated_artifacts"]
    raw_docx = ensure_inside_run_dir(run_dir / artifacts["raw_docx"], run_dir, "raw DOCX")
    final_docx = ensure_inside_run_dir(run_dir / artifacts["final_docx"], run_dir, "final DOCX")
    ris = ensure_inside_run_dir(run_dir / artifacts["ris"], run_dir, "RIS")
    if not raw_docx.exists():
        raise SystemExit(f"Missing raw revised DOCX generated in run directory: {raw_docx}")

    changed, source_count, raw_count = paragraph_differences(source_docx, raw_docx)
    allowed = set(manifest["allowed_paragraphs"])
    appended_paragraphs = set(range(source_count + 1, raw_count + 1))
    reference_integrity = set(allowed_reference_integrity_changes(source_docx, raw_docx))
    source_paragraphs = content_paragraphs(docx_roots(source_docx)[0])
    raw_paragraphs = content_paragraphs(docx_roots(raw_docx)[0])
    source_ref_start = reference_region_start(source_paragraphs)
    citation_only = {
        index
        for index in changed
        if source_ref_start
        and index < source_ref_start
        and index <= len(raw_paragraphs)
        and is_citation_only_change(source_paragraphs[index - 1], raw_paragraphs[index - 1])
    }
    reference_region = {index for index in changed if source_ref_start and index >= source_ref_start}
    unexpected = [
        index
        for index in changed
        if index not in allowed
        and index not in appended_paragraphs
        and index not in reference_integrity
        and index not in citation_only
        and index not in reference_region
    ]
    if unexpected:
        raise SystemExit(f"Unexpected paragraph changes outside Word-comment scope: {unexpected}")

    run(["python3", str(SCRIPT_DIR / "docx_reference_list_to_ris.py"), str(raw_docx), str(ris)])
    raw_reference_count = len(reference_numbers(content_paragraphs(docx_roots(raw_docx)[0])))
    allowed_reference_deficit = reference_duplicate_surplus(source_paragraphs)
    if manifest["reference_count"] and raw_reference_count + allowed_reference_deficit < manifest["reference_count"]:
        raise SystemExit("Revised raw DOCX has fewer numbered references than the source DOCX.")
    paragraph_arg = ",".join(str(index) for index in sorted(allowed)) if allowed else None
    support_cmd = [
        "python3",
        str(SCRIPT_DIR / "docx_modified_citation_support.py"),
        "--source",
        str(source_docx),
        "--revised",
        str(raw_docx),
    ]
    if paragraph_arg:
        support_cmd.extend(["--paragraphs", paragraph_arg])
    run(support_cmd)
    run(["python3", str(SCRIPT_DIR / "docx_plain_numeric_citation_check.py"), str(raw_docx)])
    run(
        [
            "python3",
            str(SCRIPT_DIR / "docx_numeric_to_endnote_temp.py"),
            str(raw_docx),
            str(final_docx),
            "--ris",
            str(ris),
            "--short-title-disambiguation",
            "--title-prefix-words",
            "6",
            "--keep-references",
        ]
    )
    run(["unzip", "-t", str(final_docx)])
    run(["python3", str(SCRIPT_DIR / "docx_word_sanity.py"), str(final_docx)])
    run(["python3", str(SCRIPT_DIR / "docx_endnote_ris_sync.py"), str(final_docx), str(ris), "--allow-title-prefix"])

    stale = {path.name: stale_marker_counts(path) for path in (raw_docx, final_docx)}
    if any(count for counts in stale.values() for count in counts.values()):
        raise SystemExit(f"Stale EndNote/comment markers remain: {stale}")

    audit = {
        "source_docx": source_docx.name,
        "source_sha256": manifest["source_sha256"],
        "changed_paragraphs": changed,
        "allowed_paragraphs": sorted(allowed),
        "appended_paragraphs": sorted(appended_paragraphs),
        "citation_only_paragraphs": sorted(citation_only),
        "reference_region_paragraphs": sorted(reference_region),
        "reference_integrity_paragraphs": sorted(reference_integrity),
        "final_docx": final_docx.name,
        "ris": ris.name,
        "reviewers": manifest.get("reviewers", {}),
        "missing_reviewer_reports": [
            report
            for report in manifest.get("reviewers", {}).get("required_reports", [])
            if not (run_dir / report).exists() or not (run_dir / report).read_text(encoding="utf-8").strip()
        ],
        "stale_marker_counts": stale,
    }
    audit_path = run_dir / "finalize_audit.json"
    audit_path.write_text(json.dumps(audit, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote final DOCX: {final_docx}")
    print(f"Wrote paired RIS: {ris}")
    print(f"Wrote audit: {audit_path}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    start_parser = subparsers.add_parser("start", help="Create a Word-doc-only revision run directory.")
    start_parser.add_argument("source_docx", help="The only external content input for the revision pass.")
    start_parser.add_argument("--output-stem", help="Output stem for generated artifacts.")
    start_parser.add_argument("--run-dir", help="Run directory to create. Defaults to OUTPUT_STEM_word_doc_only_run.")
    start_parser.set_defaults(func=start)

    finalize_parser = subparsers.add_parser("finalize", help="Validate and finish a Word-doc-only revision run.")
    finalize_parser.add_argument("manifest", help="Manifest created by the start command.")
    finalize_parser.set_defaults(func=finalize)

    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except subprocess.CalledProcessError as error:
        raise SystemExit(error.returncode)
