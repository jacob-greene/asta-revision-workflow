#!/usr/bin/env python3
"""Pandoc-centered Word revision launcher.

The current Word document remains the source of truth, but the revision
surface is a Pandoc markdown export generated inside the current run
directory. Citation membership/order comes from the recompiled Word document;
complete citation metadata is extracted from embedded EndNote fields in the
same source DOCX.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import zipfile
from dataclasses import asdict
from pathlib import Path

from citeproc_endnote_uv.docx_endnote_to_ris import export_ris as export_embedded_endnote_ris
from citeproc_endnote_uv.docx_extract_comments import extract_comments, format_markdown
from citeproc_endnote_uv.word_doc_only_revision_launcher import (
    endnote_conversion_command,
    reference_list_to_ris_command,
    stale_marker_counts,
    temporary_citation_entries,
)
from citeproc_endnote_uv.strip_docx_comments import strip_comments

SCRIPT_DIR = Path(__file__).resolve().parent
PANDOC_FROM = "docx+styles"
PANDOC_TO = "markdown+bracketed_spans+fenced_divs+link_attributes+pipe_tables+tex_math_single_backslash"


def run(command: list[str]) -> None:
    print("+ " + " ".join(command))
    subprocess.run(command, check=True)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def require_docx(path: Path) -> None:
    if path.suffix.lower() != ".docx":
        raise SystemExit(f"Source must be a .docx file: {path}")


def require_pandoc() -> None:
    if shutil.which("pandoc") is None:
        raise SystemExit("pandoc is required for pandoc-word-revision but was not found on PATH.")


def ensure_inside_run_dir(path: Path, run_dir: Path, label: str) -> Path:
    resolved = path.resolve()
    try:
        resolved.relative_to(run_dir)
    except ValueError as exc:
        raise SystemExit(f"{label} must be inside the run directory: {resolved}") from exc
    return resolved


def write_json(path: Path, data: object) -> None:
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def pandoc_docx_to_markdown(source_docx: Path, markdown: Path, media_dir: Path) -> list[str]:
    return [
        "pandoc",
        "-f",
        PANDOC_FROM,
        "-t",
        PANDOC_TO,
        "--wrap=none",
        f"--extract-media={media_dir}",
        str(source_docx),
        "-o",
        str(markdown),
    ]


def pandoc_markdown_to_docx(markdown: Path, output_docx: Path, reference_docx: Path) -> list[str]:
    return [
        "pandoc",
        "-f",
        PANDOC_TO,
        "-t",
        "docx",
        "--wrap=none",
        f"--reference-doc={reference_docx}",
        str(markdown),
        "-o",
        str(output_docx),
    ]


def start(args: argparse.Namespace) -> int:
    require_pandoc()
    source = Path(args.source_docx).resolve()
    require_docx(source)
    if not source.exists():
        raise SystemExit(f"Source DOCX does not exist: {source}")

    output_stem = args.output_stem or source.stem
    run_dir = (Path(args.run_dir) if args.run_dir else source.parent / f"{output_stem}_pandoc_revision_run").resolve()
    run_dir.mkdir(parents=True, exist_ok=True)

    source_copy = run_dir / source.name
    style_reference = run_dir / "style-reference.docx"
    markdown = run_dir / f"{output_stem}.source.md"
    revised_markdown = run_dir / f"{output_stem}.revised.md"
    comments_md = run_dir / f"{output_stem}.comments.md"
    comments_json = run_dir / f"{output_stem}.comments.json"
    media_dir = run_dir / "media"
    raw_docx = run_dir / f"{output_stem}.raw.docx"
    final_docx = run_dir / f"{output_stem}.docx"
    ris = run_dir / f"{output_stem}.ris"

    shutil.copy2(source, source_copy)
    shutil.copy2(source, style_reference)

    metadata_ris = run_dir / "citation_metadata.ris"
    metadata_audit_path = run_dir / "citation_metadata_audit.json"
    metadata_audit = export_embedded_endnote_ris(source_copy, metadata_ris)
    metadata_audit["source"] = "embedded-endnote-fields"
    metadata_ris_name = metadata_ris.name
    metadata_audit_name = metadata_audit_path.name
    if metadata_audit["missing_author_records"]:
        write_json(metadata_audit_path, metadata_audit)
        raise SystemExit(
            "Embedded EndNote metadata contains records without authors; refusing to create a truncated RIS overlay. "
            f"See {metadata_audit_path}"
        )
    if args.metadata_ris:
        metadata_source = Path(args.metadata_ris).resolve()
        if not metadata_source.exists():
            raise SystemExit(f"Metadata RIS does not exist: {metadata_source}")
        if metadata_audit["embedded_record_count"]:
            fallback = run_dir / "fallback_external_metadata.ris"
            shutil.copy2(metadata_source, fallback)
            metadata_audit["fallback_external_metadata_ris"] = fallback.name
            metadata_audit["fallback_external_metadata_used"] = False
        else:
            shutil.copy2(metadata_source, metadata_ris)
            metadata_audit["source"] = "external-metadata-ris-fallback"
            metadata_audit["fallback_external_metadata_used"] = True
    if not metadata_audit["embedded_record_count"] and not args.metadata_ris:
        raise SystemExit(
            "No embedded EndNote records were found in the source DOCX. The Pandoc workflow now derives "
            "complete citation metadata from the current Word file; provide a DOCX with EndNote fields or "
            "an explicit --metadata-ris fallback."
        )
    write_json(metadata_audit_path, metadata_audit)

    comments = extract_comments(source_copy)
    comments_md.write_text(format_markdown(comments), encoding="utf-8")
    write_json(comments_json, [asdict(comment) for comment in comments])

    run(pandoc_docx_to_markdown(source_copy, markdown, media_dir))
    if not revised_markdown.exists():
        shutil.copy2(markdown, revised_markdown)

    manifest = {
        "workflow": "pandoc-word-revision",
        "source_docx": source_copy.name,
        "source_sha256": sha256(source_copy),
        "pandoc": {
            "from": PANDOC_FROM,
            "to_markdown": PANDOC_TO,
            "reference_doc": style_reference.name,
        },
        "comments": {
            "markdown": comments_md.name,
            "json": comments_json.name,
            "count": len(comments),
        },
        "citation_policy": {
            "membership_and_order": "recompiled-current-run-docx",
            "metadata_overlay_ris": metadata_ris_name,
            "metadata_source": metadata_audit["source"],
            "metadata_audit": metadata_audit_name,
            "require_metadata_match": True,
            "new_asta_references": "recorded in asta_reference_additions.json when present",
        },
        "generated_artifacts": {
            "source_markdown": markdown.name,
            "revised_markdown": revised_markdown.name,
            "media_dir": media_dir.name,
            "raw_docx": raw_docx.name,
            "final_docx": final_docx.name,
            "ris": ris.name,
        },
    }
    manifest_path = run_dir / "manifest.json"
    write_json(manifest_path, manifest)

    print(f"Wrote run directory: {run_dir}")
    print(f"Revise markdown: {revised_markdown}")
    print(f"Finalize with: pandoc-word-revision finalize {manifest_path}")
    return 0


def finalize(args: argparse.Namespace) -> int:
    require_pandoc()
    manifest_path = Path(args.manifest).resolve()
    run_dir = manifest_path.parent
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("workflow") != "pandoc-word-revision":
        raise SystemExit("Manifest is not a pandoc-word-revision manifest.")

    source_docx = run_dir / manifest["source_docx"]
    if sha256(source_docx) != manifest["source_sha256"]:
        raise SystemExit("Source DOCX hash no longer matches the launch manifest.")

    artifacts = manifest["generated_artifacts"]
    revised_markdown = ensure_inside_run_dir(run_dir / artifacts["revised_markdown"], run_dir, "revised markdown")
    raw_docx = ensure_inside_run_dir(run_dir / artifacts["raw_docx"], run_dir, "raw DOCX")
    final_docx = ensure_inside_run_dir(run_dir / artifacts["final_docx"], run_dir, "final DOCX")
    ris = ensure_inside_run_dir(run_dir / artifacts["ris"], run_dir, "RIS")
    reference_doc = ensure_inside_run_dir(run_dir / manifest["pandoc"]["reference_doc"], run_dir, "reference DOCX")
    if not revised_markdown.exists():
        raise SystemExit(f"Missing revised markdown: {revised_markdown}")

    run(pandoc_markdown_to_docx(revised_markdown, raw_docx, reference_doc))
    stripped_raw = raw_docx.with_name(f"{raw_docx.stem}.stripped{raw_docx.suffix}")
    strip_comments(raw_docx, stripped_raw)
    stripped_raw.replace(raw_docx)

    metadata_name = manifest.get("citation_policy", {}).get("metadata_overlay_ris")
    metadata_ris = run_dir / metadata_name if metadata_name else run_dir / "citation_metadata.ris"
    if not metadata_ris.exists():
        metadata_ris = None

    run(reference_list_to_ris_command(raw_docx, ris, metadata_ris, require_metadata_match=metadata_ris is not None))
    check_ris_cmd = reference_list_to_ris_command(raw_docx, ris, metadata_ris, require_metadata_match=metadata_ris is not None)
    check_ris_cmd.append("--check")
    run(check_ris_cmd)
    run(["python3", str(SCRIPT_DIR / "docx_plain_numeric_citation_check.py"), str(raw_docx)])
    run(endnote_conversion_command(raw_docx, final_docx, ris))
    run(["unzip", "-t", str(final_docx)])
    run(["python3", str(SCRIPT_DIR / "docx_word_sanity.py"), str(final_docx)])
    run(["python3", str(SCRIPT_DIR / "docx_endnote_ris_sync.py"), str(final_docx), str(ris)])
    run(check_ris_cmd)

    repeat_docx = final_docx.with_name(f"{final_docx.stem}.determinism-check{final_docx.suffix}")
    try:
        run(endnote_conversion_command(raw_docx, repeat_docx, ris))
        primary_entries = temporary_citation_entries(final_docx)
        repeat_entries = temporary_citation_entries(repeat_docx)
        if primary_entries != repeat_entries:
            raise SystemExit("EndNote temporary citation conversion is not deterministic across repeated runs.")
    finally:
        if repeat_docx.exists():
            repeat_docx.unlink()

    stale = {path.name: stale_marker_counts(path) for path in (raw_docx, final_docx)}
    if any(count for counts in stale.values() for count in counts.values()):
        raise SystemExit(f"Stale EndNote/comment markers remain: {stale}")

    audit = {
        "workflow": "pandoc-word-revision",
        "source_docx": source_docx.name,
        "source_sha256": manifest["source_sha256"],
        "source_markdown": artifacts["source_markdown"],
        "revised_markdown": revised_markdown.name,
        "raw_docx": raw_docx.name,
        "final_docx": final_docx.name,
        "ris": ris.name,
        "citation_metadata_ris": metadata_ris.name if metadata_ris is not None else None,
        "temporary_citation_determinism_check": {
            "repeated_conversion": True,
            "temporary_citation_entries": len(primary_entries),
        },
        "stale_marker_counts": stale,
    }
    audit_path = run_dir / "finalize_audit.json"
    write_json(audit_path, audit)

    print(f"Wrote final DOCX: {final_docx}")
    print(f"Wrote paired RIS: {ris}")
    print(f"Wrote audit: {audit_path}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    start_parser = subparsers.add_parser("start", help="Create a Pandoc Word revision run directory.")
    start_parser.add_argument("source_docx")
    start_parser.add_argument("--output-stem")
    start_parser.add_argument("--run-dir")
    start_parser.add_argument(
        "--metadata-ris",
        help="Fallback complete RIS metadata overlay, used only when the source DOCX has no embedded EndNote records.",
    )
    start_parser.set_defaults(func=start)

    finalize_parser = subparsers.add_parser("finalize", help="Compile revised markdown and finalize DOCX/RIS.")
    finalize_parser.add_argument("manifest")
    finalize_parser.set_defaults(func=finalize)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
