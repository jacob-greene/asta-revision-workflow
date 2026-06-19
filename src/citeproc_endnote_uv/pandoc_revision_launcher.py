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
import os
import re
import shlex
import shutil
import subprocess
import sys
import time
import zipfile
from contextlib import contextmanager
from dataclasses import asdict
from pathlib import Path
from xml.etree import ElementTree as ET

from citeproc_endnote_uv.docx_endnote_to_ris import export_ris as export_embedded_endnote_ris
from citeproc_endnote_uv.docx_extract_comments import extract_comments, format_markdown
from citeproc_endnote_uv.strip_docx_comments import strip_comments

SCRIPT_DIR = Path(__file__).resolve().parent
W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"
PANDOC_FROM = "docx+styles"
PANDOC_TO = "markdown+bracketed_spans+fenced_divs+link_attributes+pipe_tables+tex_math_single_backslash"
COMMENT_PARTS = {
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsExtensible.xml",
    "word/commentsIds.xml",
}
AGENT_WORKFLOW_PASSES = [
    {
        "name": "revision_implementation",
        "report": "revision_implementation_report.md",
        "required_checks": [
            "revisions_applied",
            "comment_scope_preserved",
            "source_docx_only",
            "draft_scientific_paper_skill_used",
            "edit_scientific_prose_skill_used",
        ],
        "required_skills": ["draft-scientific-paper", "edit-scientific-prose"],
        "instruction": (
            "Implement comment-scoped revisions from the current Word-derived markdown only. Do not reintroduce "
            "deleted text from older drafts. Prefer precise replacement of commented paragraphs or immediately "
            "adjacent paragraphs only when required by the comment. Do not add unsupported knowledge claims. "
            "If a retained claim needs evidence beyond adjacent citations, add a required entry to "
            "`agent_workflow/asta_requests.json`. Because nested shell access may be unavailable, return a "
            "fenced JSON payload in this report with `markdown_replacements`, each containing exact `old` and "
            "`new` strings to apply to the revised markdown. Do not return an empty `markdown_replacements` "
            "list unless every actionable comment is already addressed in the current revised markdown or every "
            "remaining comment is impossible to revise from the current DOCX. At minimum, implement "
            "style, clarity, typo, structure, and redundancy edits for comments whose core evidence already "
            "appears in adjacent citations; defer only the specific unsupported knowledge clauses to Asta. "
            "Do not expand section headings into full-sentence mechanistic claims or add citations to headings; "
            "keep headings short unless the comment explicitly requests a heading change."
        ),
    },
    {
        "name": "comment_interpretation_and_revision_planning",
        "report": "comment_plan_report.md",
        "required_checks": [
            "comments_addressed",
            "revision_scope_defined",
            "source_docx_only",
            "draft_scientific_paper_skill_used",
        ],
        "required_skills": ["draft-scientific-paper"],
        "instruction": (
            "Read the run-local source markdown, revised markdown, comments markdown/json, and manifest. "
            "Produce a comment-keyed plan, current outline, exact allowed revision scope, and any justified "
            "adjacent-paragraph exceptions. Treat a comment as addressed when unsupported requested specificity "
            "was explicitly narrowed or omitted because adjacent citations or resolved Asta evidence did not "
            "support it; do not require adding unsupported examples merely because a reviewer comment suggested "
            "them."
        ),
    },
    {
        "name": "evidence_and_specificity",
        "report": "evidence_specificity_report.md",
        "required_checks": [
            "modified_claims_citation_checked",
            "unsupported_claims_resolved",
            "source_docx_only",
            "draft_scientific_paper_skill_used",
        ],
        "required_skills": ["draft-scientific-paper"],
        "instruction": (
            "Check each modified claim for same-sentence or adjacent citation support. If nearby existing "
            "citations do not support the claim, prioritize writing a required request to "
            "`agent_workflow/asta_requests.json` so the launcher can query Asta before finalize. "
            "Soften or remove the claim only when the claim should not be retained even with new evidence."
        ),
    },
    {
        "name": "rigor_critique",
        "report": "rigor_critique_report.md",
        "required_checks": [
            "rigor_approved",
            "new_knowledge_claims_skeptically_reviewed",
            "uncommented_changes_reviewed",
            "draft_scientific_paper_skill_used",
        ],
        "required_skills": ["draft-scientific-paper"],
        "instruction": (
            "Be highly skeptical of new knowledge claims, broad causal language, conserved/universal claims, "
            "and accidental edits to uncommented text. Approve only narrow claims with explicit support."
        ),
    },
    {
        "name": "tone_and_concision",
        "report": "tone_concision_report.md",
        "required_checks": [
            "tone_reviewed",
            "redundancy_checked",
            "comment_scope_preserved",
            "edit_scientific_prose_skill_used",
        ],
        "required_skills": ["edit-scientific-prose"],
        "instruction": (
            "Review topic sentences, paragraph flow, concision, and thesis tone. Flag restatement of nearby "
            "material and tone drift."
        ),
    },
]
def run(command: list[str]) -> None:
    print("+ " + " ".join(command))
    result = subprocess.run(command)
    if result.returncode != 0:
        raise SystemExit(f"Command failed with exit code {result.returncode}: {' '.join(command)}")


@contextmanager
def timed_step(profile: list[dict[str, object]] | None, name: str):
    start = time.perf_counter()
    yield
    if profile is not None:
        profile.append({"step": name, "seconds": round(time.perf_counter() - start, 4)})


def endnote_conversion_command(raw_docx: Path, output_docx: Path, ris: Path) -> list[str]:
    return [
        sys.executable,
        "-m",
        "citeproc_endnote_uv.docx_numeric_to_endnote_temp",
        str(raw_docx),
        str(output_docx),
        "--ris",
        str(ris),
        "--keep-references",
    ]


def reference_list_to_ris_command(
    source_docx: Path, ris: Path, metadata_ris: Path | None = None, require_metadata_match: bool = False
) -> list[str]:
    command = [
        sys.executable,
        "-m",
        "citeproc_endnote_uv.docx_reference_list_to_ris",
        str(source_docx),
        str(ris),
    ]
    if metadata_ris is not None:
        command.extend(["--metadata-ris", str(metadata_ris)])
    if require_metadata_match:
        command.append("--require-metadata-match")
    return command


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def text_of(elem: ET.Element) -> str:
    return "".join(t.text or "" for t in elem.findall(f".//{W}t"))


def docx_visible_text(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    return text_of(root)


def temporary_citation_entries(path: Path) -> list[str]:
    entries: list[str] = []
    for citation in re.findall(r"\{([^{}]+)\}", docx_visible_text(path)):
        entries.extend(re.sub(r"\s+", " ", part).strip() for part in citation.split(";") if part.strip())
    return entries


def stale_marker_counts(docx: Path) -> dict[str, int]:
    with zipfile.ZipFile(docx) as zf:
        data = "\n".join(zf.namelist()) + "\n"
        for name in zf.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                data += zf.read(name).decode("utf-8", "ignore")
    tokens = [
        "ADDIN EN.",
        "ADDIN EN.CITE",
        "ADDIN EN.REFLIST",
        "EN.CWYW",
        "commentRangeStart",
        "commentRangeEnd",
        "commentReference",
    ]
    counts = {token: data.count(token) for token in tokens}
    counts["comment_parts"] = sum(1 for part in COMMENT_PARTS if part in data)
    return counts


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


def relative_to_run(path: Path, run_dir: Path) -> str:
    return str(path.relative_to(run_dir))


def approx_text_tokens(text: str) -> int:
    return (len(text) + 3) // 4


def artifact_profile(path: Path, run_dir: Path, name: str) -> dict[str, object]:
    size = path.stat().st_size if path.exists() else 0
    suffix = path.suffix.lower()
    tokens = 0
    if path.exists() and suffix not in {".docx", ".png", ".jpg", ".jpeg", ".gif", ".pdf"}:
        tokens = approx_text_tokens(path.read_text(encoding="utf-8", errors="ignore"))
    return {
        "artifact": name,
        "path": relative_to_run(path, run_dir),
        "bytes": size,
        "approx_tokens": tokens,
    }


def comment_scope_from_json(comments_json: Path, fallback_markdown: Path) -> str:
    comments = json.loads(comments_json.read_text(encoding="utf-8"))
    if not comments:
        return fallback_markdown.read_text(encoding="utf-8")
    chunks: list[str] = []
    for comment in comments:
        chunks.append(
            "\n".join(
                [
                    f"## Comment {comment.get('comment_id', '')}",
                    "",
                    f"Paragraph: {comment.get('paragraph_index', '')}",
                    "",
                    "Comment:",
                    str(comment.get("comment_text", "")).strip(),
                    "",
                    "Anchored text:",
                    str(comment.get("anchored_text", "")).strip(),
                    "",
                    "Source paragraph:",
                    str(comment.get("paragraph_text", "")).strip(),
                ]
            )
        )
    return "\n\n".join(chunks).strip() + "\n"


def citation_numbers_from_text(text: str) -> list[int]:
    numbers = sorted({int(value) for value in re.findall(r"\b\d+\b", text)})
    return [number for number in numbers if number > 0]


def write_scoped_citation_metadata(path: Path, citation_metadata: Path, scoped_text: str) -> None:
    numbers = citation_numbers_from_text(scoped_text)
    if not citation_metadata.exists():
        path.write_text("No citation metadata RIS was available.\n", encoding="utf-8")
        return
    records = parse_ris_records(citation_metadata.read_text(encoding="utf-8", errors="ignore"))
    rows = [
        "# Scoped Citation Metadata",
        "",
        "Citation records below are limited to citation numbers visible in Word-commented passages.",
        "",
    ]
    for number in numbers:
        if number < 1 or number > len(records):
            continue
        record = records[number - 1]
        title = (record.get("TI") or [""])[0]
        year = (record.get("PY") or [""])[0]
        journal = (record.get("JO") or [""])[0]
        doi = (record.get("DO") or [""])[0]
        authors = "; ".join(record.get("AU", [])[:6])
        if len(record.get("AU", [])) > 6:
            authors += "; et al."
        rows.extend(
            [
                f"## Reference {number}",
                "",
                f"Title: {title}",
                f"Year: {year}",
                f"Authors: {authors}",
                f"Journal: {journal}",
                f"DOI: {doi}",
                "",
            ]
        )
    path.write_text("\n".join(rows).rstrip() + "\n", encoding="utf-8")


def write_agent_inputs(run_dir: Path, manifest: dict, source_markdown: Path, revised_markdown: Path) -> dict[str, object]:
    inputs_dir = run_dir / "agent_inputs"
    inputs_dir.mkdir(exist_ok=True)
    comments_md = run_dir / manifest["comments"]["markdown"]
    comments_json = run_dir / manifest["comments"]["json"]
    citation_metadata = run_dir / manifest["citation_policy"]["metadata_overlay_ris"]
    citation_summary = inputs_dir / "citation_metadata_scoped.md"
    source_scoped = inputs_dir / "comment_scoped_source.md"
    revised_scoped = inputs_dir / "comment_scoped_revised.md"
    agent_input_manifest = inputs_dir / "agent_input_manifest.json"

    scoped_text = comment_scope_from_json(comments_json, source_markdown)
    write_scoped_citation_metadata(citation_summary, citation_metadata, scoped_text)
    source_scoped.write_text(scoped_text, encoding="utf-8")
    revised_scoped.write_text(
        scoped_text,
        encoding="utf-8",
    )

    common = [
        relative_to_run(agent_input_manifest, run_dir),
        relative_to_run(comments_md, run_dir),
        relative_to_run(source_scoped, run_dir),
        relative_to_run(revised_scoped, run_dir),
        "manifest.json",
    ]
    passes = {
        "revision_implementation": {
            "recommended_inputs": common + [relative_to_run(comments_json, run_dir)],
            "avoid_by_default": [relative_to_run(citation_metadata, run_dir)],
        },
        "comment_interpretation_and_revision_planning": {
            "recommended_inputs": common + [relative_to_run(comments_json, run_dir)],
            "avoid_by_default": [relative_to_run(citation_metadata, run_dir)],
        },
        "evidence_and_specificity": {
            "recommended_inputs": common + [relative_to_run(comments_json, run_dir), relative_to_run(citation_summary, run_dir)],
            "avoid_by_default": [relative_to_run(citation_metadata, run_dir)],
        },
        "rigor_critique": {
            "recommended_inputs": common,
            "avoid_by_default": [relative_to_run(citation_metadata, run_dir)],
        },
        "tone_and_concision": {
            "recommended_inputs": common,
            "avoid_by_default": [relative_to_run(citation_metadata, run_dir), relative_to_run(comments_json, run_dir)],
        },
    }
    data = {
        "purpose": "Prefer these scoped inputs for agent passes; load full markdown or citation_metadata.ris only when needed.",
        "scoped_source_markdown": relative_to_run(source_scoped, run_dir),
        "scoped_revised_markdown": relative_to_run(revised_scoped, run_dir),
        "full_source_markdown": manifest["generated_artifacts"]["source_markdown"],
        "full_revised_markdown": manifest["generated_artifacts"]["revised_markdown"],
        "comments_markdown": manifest["comments"]["markdown"],
        "comments_json": manifest["comments"]["json"],
        "citation_metadata_ris": manifest["citation_policy"]["metadata_overlay_ris"],
        "scoped_citation_metadata": relative_to_run(citation_summary, run_dir),
        "passes": passes,
    }
    write_json(agent_input_manifest, data)
    return {
        "directory": relative_to_run(inputs_dir, run_dir),
        "manifest": relative_to_run(agent_input_manifest, run_dir),
        "comment_scoped_source_markdown": relative_to_run(source_scoped, run_dir),
        "comment_scoped_revised_markdown": relative_to_run(revised_scoped, run_dir),
        "pass_input_policy": passes,
    }


def write_launcher_profile(
    path: Path,
    run_dir: Path,
    source: Path,
    manifest: dict,
    steps: list[dict[str, object]],
    metadata_audit: dict[str, object],
    comment_count: int,
) -> None:
    artifacts = manifest["generated_artifacts"]
    citation_policy = manifest["citation_policy"]
    comments = manifest["comments"]
    tracked_paths = {
        "source_docx_copy": run_dir / manifest["source_docx"],
        "style_reference_docx": run_dir / manifest["pandoc"]["reference_doc"],
        "source_markdown": run_dir / artifacts["source_markdown"],
        "revised_markdown": run_dir / artifacts["revised_markdown"],
        "comment_scoped_source_markdown": run_dir / manifest["agent_inputs"]["comment_scoped_source_markdown"],
        "comment_scoped_revised_markdown": run_dir / manifest["agent_inputs"]["comment_scoped_revised_markdown"],
        "comments_markdown": run_dir / comments["markdown"],
        "comments_json": run_dir / comments["json"],
        "citation_metadata_ris": run_dir / citation_policy["metadata_overlay_ris"],
        "citation_metadata_audit": run_dir / citation_policy["metadata_audit"],
        "agent_input_manifest": run_dir / manifest["agent_inputs"]["manifest"],
        "agent_audit_template": run_dir / manifest["agent_workflow"]["audit_template"],
        "manifest": run_dir / "manifest.json",
    }
    for index, task in enumerate(manifest["agent_workflow"]["task_files"], start=1):
        tracked_paths[f"agent_task_{index}"] = run_dir / task

    artifact_rows = [
        artifact_profile(artifact_path, run_dir, name)
        for name, artifact_path in tracked_paths.items()
        if artifact_path.exists()
    ]
    row_by_name = {row["artifact"]: row for row in artifact_rows}
    full_context_names = {
        "manifest",
        "source_markdown",
        "revised_markdown",
        "comments_markdown",
        "comments_json",
        "citation_metadata_ris",
    }
    scoped_common_names = {
        "manifest",
        "comment_scoped_source_markdown",
        "comment_scoped_revised_markdown",
        "comments_markdown",
        "agent_input_manifest",
    }
    full_context_tokens = sum(int(row_by_name[name]["approx_tokens"]) for name in full_context_names if name in row_by_name)
    scoped_common_tokens = sum(int(row_by_name[name]["approx_tokens"]) for name in scoped_common_names if name in row_by_name)

    pass_estimates = []
    for workflow_pass in AGENT_WORKFLOW_PASSES:
        input_policy = manifest["agent_inputs"]["pass_input_policy"][workflow_pass["name"]]
        tokens = 0
        input_rows = []
        for input_path in input_policy["recommended_inputs"]:
            matching = next((row for row in artifact_rows if row["path"] == input_path), None)
            if matching is not None:
                tokens += int(matching["approx_tokens"])
                input_rows.append(matching)
        task_path = f"agent_workflow/tasks/{workflow_pass['name']}.md"
        task_row = next((row for row in artifact_rows if row["path"] == task_path), None)
        if task_row is not None:
            tokens += int(task_row["approx_tokens"])
        pass_estimates.append(
            {
                "pass": workflow_pass["name"],
                "recommended_input_tokens": tokens,
                "full_context_tokens_if_loaded": full_context_tokens + (int(task_row["approx_tokens"]) if task_row else 0),
                "estimated_token_savings": max(
                    0,
                    (full_context_tokens + (int(task_row["approx_tokens"]) if task_row else 0)) - tokens,
                ),
            }
        )

    profile = {
        "run_dir": str(run_dir),
        "source_docx": str(source),
        "steps": steps,
        "total_profiled_seconds": round(sum(float(item["seconds"]) for item in steps), 4),
        "embedded_record_count": metadata_audit.get("embedded_record_count"),
        "comment_count": comment_count,
        "artifact_rows": artifact_rows,
        "generated_text_artifact_tokens": sum(int(row["approx_tokens"]) for row in artifact_rows),
        "full_context_agent_tokens": full_context_tokens,
        "scoped_common_agent_tokens": scoped_common_tokens,
        "agent_pass_estimates": pass_estimates,
        "four_pass_full_context_total_tokens": sum(int(item["full_context_tokens_if_loaded"]) for item in pass_estimates),
        "four_pass_recommended_total_tokens": sum(int(item["recommended_input_tokens"]) for item in pass_estimates),
        "four_pass_estimated_token_savings": sum(int(item["estimated_token_savings"]) for item in pass_estimates),
        "billing_note": "Launcher uses local tools only and makes no LLM/API calls; token values are approximate prompt-size estimates.",
    }
    write_json(path, profile)


def write_agent_workflow_tasks(run_dir: Path, manifest: dict) -> dict[str, object]:
    workflow_dir = run_dir / "agent_workflow"
    tasks_dir = workflow_dir / "tasks"
    reports_dir = workflow_dir / "reports"
    tasks_dir.mkdir(parents=True, exist_ok=True)
    reports_dir.mkdir(parents=True, exist_ok=True)

    task_files: list[str] = []
    required_reports: list[str] = []
    passes: list[dict[str, object]] = []
    artifacts = manifest["generated_artifacts"]
    for workflow_pass in AGENT_WORKFLOW_PASSES:
        input_policy = manifest.get("agent_inputs", {}).get("pass_input_policy", {}).get(workflow_pass["name"], {})
        task_path = tasks_dir / f"{workflow_pass['name']}.md"
        report_path = reports_dir / str(workflow_pass["report"])
        task_path.write_text(
            "\n".join(
                [
                    f"# {str(workflow_pass['name']).replace('_', ' ').title()}",
                    "",
                    str(workflow_pass["instruction"]),
                    "",
                    "## Required Inputs",
                    f"- Manifest: `manifest.json`",
                    f"- Source DOCX: `{manifest['source_docx']}`",
                    f"- Full source markdown, only if needed: `{artifacts['source_markdown']}`",
                    f"- Full revised markdown, only if needed: `{artifacts['revised_markdown']}`",
                    f"- Comments markdown: `{manifest['comments']['markdown']}`",
                    f"- Comments JSON: `{manifest['comments']['json']}`",
                    f"- Citation metadata RIS, only if needed: `{manifest['citation_policy']['metadata_overlay_ris']}`",
                    "- Asta request ledger: `agent_workflow/asta_requests.json`",
                    "",
                    "## Recommended Minimal Inputs",
                    *[f"- `{path}`" for path in input_policy.get("recommended_inputs", [])],
                    "",
                    "## Avoid Loading By Default",
                    *[f"- `{path}`" for path in input_policy.get("avoid_by_default", [])],
                    "",
                    "## Required Checks",
                    *[f"- `{check}`" for check in workflow_pass["required_checks"]],
                    "",
                    "## Required Codex Skills",
                    *[f"- `{skill}`" for skill in workflow_pass.get("required_skills", [])],
                    "",
                    "Write the report to:",
                    f"`{relative_to_run(report_path, run_dir)}`",
                    "",
                ]
            ),
            encoding="utf-8",
        )
        task_files.append(relative_to_run(task_path, run_dir))
        report_rel = relative_to_run(report_path, run_dir)
        required_reports.append(report_rel)
        passes.append(
            {
                "name": workflow_pass["name"],
                "required_checks": workflow_pass["required_checks"],
                "required_skills": workflow_pass.get("required_skills", []),
                "report": report_rel,
            }
        )

    audit_template = {
        "workflow": "pandoc-word-revision-agent-workflow",
        "source_sha256": manifest["source_sha256"],
        "revised_markdown": artifacts["revised_markdown"],
        "revised_markdown_sha256": "<fill after final edits>",
        "passes": [
            {
                "name": workflow_pass["name"],
                "status": "pending",
                "report": f"agent_workflow/reports/{workflow_pass['report']}",
                "checks": {check: False for check in workflow_pass["required_checks"]},
                "required_skills": workflow_pass.get("required_skills", []),
            }
            for workflow_pass in AGENT_WORKFLOW_PASSES
        ],
        "overall": {
            "all_comments_addressed": False,
            "modified_claims_have_adjacent_citation_or_resolution": False,
            "uncommented_changes_justified": False,
            "citation_integrity_reviewed": False,
            "ready_for_finalize": False,
        },
    }
    template_path = workflow_dir / "agent_workflow_audit.template.json"
    write_json(template_path, audit_template)
    return {
        "required": True,
        "tasks_dir": relative_to_run(tasks_dir, run_dir),
        "task_files": task_files,
        "required_reports": required_reports,
        "audit_template": relative_to_run(template_path, run_dir),
        "audit_file": "agent_workflow/agent_workflow_audit.json",
        "asta_requests": "agent_workflow/asta_requests.json",
        "asta_resolutions_dir": "agent_workflow/asta",
        "required_passes": passes,
    }


def write_asta_request_template(run_dir: Path, workflow: dict[str, object]) -> None:
    requests_path = run_dir / str(workflow["asta_requests"])
    requests_path.parent.mkdir(parents=True, exist_ok=True)
    if requests_path.exists():
        return
    write_json(
        requests_path,
        {
            "version": 1,
            "purpose": (
                "Evidence reviewers add required requests here when a modified claim cannot be supported by "
                "adjacent citations from the current DOCX. Asta requests are the default resolution for "
                "claims that should be retained. Finalize resolves pending required requests with "
                "--asta-command or PANDOC_REVISION_ASTA_COMMAND, validates complete RIS, and fails if unresolved."
            ),
            "requests": [],
        },
    )


def parse_ris_records(text: str) -> list[dict[str, list[str]]]:
    records: list[dict[str, list[str]]] = []
    current: dict[str, list[str]] | None = None
    for raw_line in text.splitlines():
        line = raw_line.rstrip()
        if not line:
            continue
        if len(line) < 5 or line[2:5] != "  -":
            continue
        tag = line[:2]
        value = line[6:].strip() if len(line) > 5 else ""
        if tag == "TY":
            current = {"TY": [value]}
            continue
        if current is None:
            continue
        if tag == "ER":
            records.append(current)
            current = None
            continue
        current.setdefault(tag, []).append(value)
    return records


def validate_complete_ris(path: Path, label: str) -> int:
    if not path.exists() or not path.read_text(encoding="utf-8", errors="ignore").strip():
        raise SystemExit(f"{label} did not produce a non-empty RIS file: {path}")
    records = parse_ris_records(path.read_text(encoding="utf-8", errors="ignore"))
    if not records:
        raise SystemExit(f"{label} RIS contains no complete records ending in ER: {path}")
    incomplete: list[str] = []
    for index, record in enumerate(records, start=1):
        missing = [tag for tag in ("TY", "TI", "AU", "PY") if not record.get(tag)]
        if missing:
            identifier = (record.get("ID") or record.get("TI") or [f"record {index}"])[0]
            incomplete.append(f"{identifier}: missing {', '.join(missing)}")
    if incomplete:
        raise SystemExit(f"{label} RIS has incomplete records:\n" + "\n".join(incomplete))
    return len(records)


def normalized_request_id(raw: object, index: int) -> str:
    text = re.sub(r"[^A-Za-z0-9_.-]+", "-", str(raw or f"request-{index}")).strip("-")
    return text or f"request-{index}"


def asta_command_parts(command: str, request_json: Path, output_json: Path, output_ris: Path, run_dir: Path, request_id: str) -> list[str]:
    replacements = {
        "request_json": str(request_json),
        "output_json": str(output_json),
        "output_ris": str(output_ris),
        "run_dir": str(run_dir),
        "request_id": request_id,
    }
    if "{" in command and "}" in command:
        return [part.format(**replacements) for part in shlex.split(command)]
    return [
        *shlex.split(command),
        "--request",
        str(request_json),
        "--output",
        str(output_json),
        "--ris",
        str(output_ris),
    ]


def workflow_agent_command_parts(command: str, run_dir: Path, manifest_path: Path, manifest: dict) -> list[str]:
    workflow = manifest.get("agent_workflow", {})
    artifacts = manifest.get("generated_artifacts", {})
    replacements = {
        "run_dir": str(run_dir),
        "manifest": str(manifest_path),
        "source_docx": str(run_dir / manifest["source_docx"]),
        "source_markdown": str(run_dir / artifacts["source_markdown"]),
        "revised_markdown": str(run_dir / artifacts["revised_markdown"]),
        "comments_markdown": str(run_dir / manifest["comments"]["markdown"]),
        "comments_json": str(run_dir / manifest["comments"]["json"]),
        "audit_file": str(run_dir / workflow.get("audit_file", "agent_workflow/agent_workflow_audit.json")),
        "asta_requests": str(run_dir / workflow.get("asta_requests", "agent_workflow/asta_requests.json")),
    }
    if "{" in command and "}" in command:
        return [part.format(**replacements) for part in shlex.split(command)]
    return [
        *shlex.split(command),
        "--manifest",
        str(manifest_path),
        "--run-dir",
        str(run_dir),
    ]


def run_agent_workflow_command(run_dir: Path, manifest_path: Path, command: str | None) -> None:
    resolver = command or os.environ.get("PANDOC_REVISION_AGENT_COMMAND")
    if not resolver:
        raise SystemExit(
            "The complete launcher requires an agent workflow command. "
            "Set PANDOC_REVISION_AGENT_COMMAND or pass --agent-command so step 2 can revise, review, "
            "write agent reports, update agent_workflow/asta_requests.json when needed, and write "
            "agent_workflow/agent_workflow_audit.json. The launcher no longer falls back to a "
            "manual handoff."
        )
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    run(workflow_agent_command_parts(resolver, run_dir, manifest_path, manifest))


def run_asta_preflight_command(run_dir: Path, manifest: dict, command: str | None) -> None:
    resolver = command or os.environ.get("PANDOC_REVISION_ASTA_COMMAND")
    if not resolver:
        return

    workflow = manifest.get("agent_workflow", {})
    asta_dir = ensure_inside_run_dir(
        run_dir / str(workflow.get("asta_resolutions_dir", "agent_workflow/asta")),
        run_dir,
        "Asta output directory",
    )
    preflight_dir = asta_dir / "preflight"
    preflight_dir.mkdir(parents=True, exist_ok=True)
    request_json = preflight_dir / "request.json"
    output_json = preflight_dir / "response.json"
    output_ris = preflight_dir / "response.ris"
    write_json(
        request_json,
        {
            "preflight": True,
            "required": True,
            "purpose": "Validate Asta CLI availability and authentication before launching revision agents.",
        },
    )
    run(asta_command_parts(resolver, request_json, output_json, output_ris, run_dir, "preflight"))
    validate_complete_ris(output_ris, "Asta preflight")


def report_requests_asta(run_dir: Path, workflow: dict[str, object]) -> bool:
    patterns = [
        re.compile(r"\b(?:needs?|requires?|required)\s+(?:an\s+)?Asta\s+requery\b", re.IGNORECASE),
        re.compile(r"\bAsta\s+requery\s+(?:needed|required)\b", re.IGNORECASE),
        re.compile(r"\bquery\s+Asta\b", re.IGNORECASE),
    ]
    for report in workflow.get("required_reports", []):
        path = run_dir / str(report)
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8", errors="ignore")
        if any(pattern.search(text) for pattern in patterns):
            return True
    return False


def resolve_asta_requests(run_dir: Path, manifest: dict, command: str | None) -> Path | None:
    workflow = manifest.get("agent_workflow", {})
    requests_rel = workflow.get("asta_requests", "agent_workflow/asta_requests.json")
    requests_path = ensure_inside_run_dir(run_dir / str(requests_rel), run_dir, "Asta request ledger")
    if not requests_path.exists():
        if report_requests_asta(run_dir, workflow):
            raise SystemExit(
                "Agent reports require Asta requery, but the Asta request ledger is missing. "
                f"Expected {requests_path}."
            )
        return None

    ledger = json.loads(requests_path.read_text(encoding="utf-8"))
    requests = ledger.get("requests", [])
    if not isinstance(requests, list):
        raise SystemExit(f"Asta request ledger must contain a list at `requests`: {requests_path}")

    required = [item for item in requests if isinstance(item, dict) and item.get("required", True)]
    pending = [item for item in required if item.get("status", "pending") not in {"resolved", "not_needed"}]
    if not pending and report_requests_asta(run_dir, workflow):
        resolved = [item for item in required if item.get("status") == "resolved"]
        if not resolved:
            raise SystemExit(
                "Agent reports mention Asta requery but no resolved Asta requests are recorded. "
                f"Add requests to {requests_path} or remove/soften unsupported claims."
            )
    if not pending:
        return combine_asta_ris(run_dir, manifest, requests_path, ledger)

    resolver = command or os.environ.get("PANDOC_REVISION_ASTA_COMMAND")
    if not resolver:
        ids = ", ".join(normalized_request_id(item.get("id"), index) for index, item in enumerate(pending, start=1))
        raise SystemExit(
            "Asta evidence is required before finalize, but no resolver command is configured. "
            "Set PANDOC_REVISION_ASTA_COMMAND or pass --asta-command. Pending request ids: " + ids
        )

    asta_dir = ensure_inside_run_dir(run_dir / str(workflow.get("asta_resolutions_dir", "agent_workflow/asta")), run_dir, "Asta output directory")
    request_dir = asta_dir / "requests"
    response_dir = asta_dir / "responses"
    request_dir.mkdir(parents=True, exist_ok=True)
    response_dir.mkdir(parents=True, exist_ok=True)

    for index, item in enumerate(pending, start=1):
        request_id = normalized_request_id(item.get("id"), index)
        request_json = request_dir / f"{request_id}.json"
        output_json = response_dir / f"{request_id}.json"
        output_ris = response_dir / f"{request_id}.ris"
        write_json(request_json, item)
        run(asta_command_parts(resolver, request_json, output_json, output_ris, run_dir, request_id))
        record_count = validate_complete_ris(output_ris, f"Asta request {request_id}")
        item["status"] = "resolved"
        item["resolution"] = {
            **(item.get("resolution") if isinstance(item.get("resolution"), dict) else {}),
            "request_json": relative_to_run(request_json, run_dir),
            "response_json": relative_to_run(output_json, run_dir) if output_json.exists() else None,
            "ris_file": relative_to_run(output_ris, run_dir),
            "ris_record_count": record_count,
        }
    write_json(requests_path, ledger)
    return combine_asta_ris(run_dir, manifest, requests_path, ledger)


def combine_asta_ris(run_dir: Path, manifest: dict, requests_path: Path, ledger: dict) -> Path | None:
    metadata_name = manifest.get("citation_policy", {}).get("metadata_overlay_ris")
    if not metadata_name:
        return None
    base_ris = run_dir / str(metadata_name)
    if not base_ris.exists():
        return None

    additions: list[Path] = []
    for item in ledger.get("requests", []):
        if not isinstance(item, dict) or item.get("status") != "resolved":
            continue
        resolution = item.get("resolution") if isinstance(item.get("resolution"), dict) else {}
        ris_file = resolution.get("ris_file")
        if ris_file:
            additions.append(ensure_inside_run_dir(run_dir / str(ris_file), run_dir, "Asta RIS addition"))

    if not additions:
        return None

    combined = run_dir / "citation_metadata.with_asta.ris"
    chunks = [base_ris.read_text(encoding="utf-8", errors="ignore").rstrip(), ""]
    total_records = 0
    for path in additions:
        total_records += validate_complete_ris(path, f"Asta RIS addition {relative_to_run(path, run_dir)}")
        chunks.append(path.read_text(encoding="utf-8", errors="ignore").strip())
        chunks.append("")
    combined.write_text("\n".join(chunks).strip() + "\n", encoding="utf-8")
    write_json(
        run_dir / "asta_reference_additions.json",
        {
            "requests": relative_to_run(requests_path, run_dir),
            "combined_metadata_ris": combined.name,
            "addition_files": [relative_to_run(path, run_dir) for path in additions],
            "addition_record_count": total_records,
        },
    )
    return combined


def validate_agent_workflow(run_dir: Path, manifest: dict, revised_markdown: Path) -> dict:
    workflow = manifest.get("agent_workflow", {})
    if not workflow.get("required", False):
        raise SystemExit("Manifest does not require the agent workflow; rerun `pandoc-word-revision start`.")

    audit_path = ensure_inside_run_dir(run_dir / workflow.get("audit_file", ""), run_dir, "agent workflow audit")
    if not audit_path.exists():
        raise SystemExit(
            "Missing required agent workflow audit. Complete the four-pass agent workflow and write "
            f"{audit_path} before finalize."
        )
    audit = json.loads(audit_path.read_text(encoding="utf-8"))
    if audit.get("workflow") != "pandoc-word-revision-agent-workflow":
        raise SystemExit("Agent workflow audit has the wrong workflow identifier.")
    if audit.get("source_sha256") != manifest["source_sha256"]:
        raise SystemExit("Agent workflow audit source hash does not match the launch manifest.")
    if audit.get("revised_markdown") != manifest["generated_artifacts"]["revised_markdown"]:
        raise SystemExit("Agent workflow audit does not name the manifest revised markdown.")
    if audit.get("revised_markdown_sha256") != sha256(revised_markdown):
        raise SystemExit("Agent workflow audit hash does not match the revised markdown being finalized.")

    passes_by_name = {item.get("name"): item for item in audit.get("passes", []) if isinstance(item, dict)}
    missing: list[str] = []
    incomplete: list[str] = []
    for required in workflow.get("required_passes", []):
        name = required["name"]
        item = passes_by_name.get(name)
        if item is None:
            missing.append(name)
            continue
        if item.get("status") != "completed":
            incomplete.append(f"{name}: status is not completed")
        report = item.get("report") or required.get("report")
        report_path = ensure_inside_run_dir(run_dir / report, run_dir, f"{name} report")
        if not report_path.exists() or not report_path.read_text(encoding="utf-8").strip():
            incomplete.append(f"{name}: missing or empty report {report}")
        checks = item.get("checks", {})
        for check in required.get("required_checks", []):
            if not checks.get(check):
                incomplete.append(f"{name}: required check `{check}` is not true")
    if missing or incomplete:
        detail = "\n".join([*(f"missing pass: {name}" for name in missing), *incomplete])
        raise SystemExit(f"Agent workflow is incomplete:\n{detail}")

    overall = audit.get("overall", {})
    required_overall = [
        "all_comments_addressed",
        "modified_claims_have_adjacent_citation_or_resolution",
        "uncommented_changes_justified",
        "citation_integrity_reviewed",
        "ready_for_finalize",
    ]
    failed_overall = [key for key in required_overall if not overall.get(key)]
    if failed_overall:
        raise SystemExit(f"Agent workflow audit is not ready for finalize; false checks: {failed_overall}")
    return audit


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
    profile_steps: list[dict[str, object]] = []
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
    launcher_profile = run_dir / "launcher_profile.json"

    with timed_step(profile_steps, "copy_source_and_style_reference"):
        shutil.copy2(source, source_copy)
        shutil.copy2(source, style_reference)

    metadata_ris = run_dir / "citation_metadata.ris"
    metadata_audit_path = run_dir / "citation_metadata_audit.json"
    with timed_step(profile_steps, "extract_embedded_endnote_metadata_to_ris"):
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
            with timed_step(profile_steps, "copy_unused_external_metadata_fallback"):
                shutil.copy2(metadata_source, fallback)
            metadata_audit["fallback_external_metadata_ris"] = fallback.name
            metadata_audit["fallback_external_metadata_used"] = False
        else:
            with timed_step(profile_steps, "copy_external_metadata_fallback"):
                shutil.copy2(metadata_source, metadata_ris)
            metadata_audit["source"] = "external-metadata-ris-fallback"
            metadata_audit["fallback_external_metadata_used"] = True
    if not metadata_audit["embedded_record_count"] and not args.metadata_ris:
        raise SystemExit(
            "No embedded EndNote records were found in the source DOCX. The Pandoc workflow now derives "
            "complete citation metadata from the current Word file; provide a DOCX with EndNote fields or "
            "an explicit --metadata-ris fallback."
        )
    with timed_step(profile_steps, "write_citation_metadata_audit"):
        write_json(metadata_audit_path, metadata_audit)

    with timed_step(profile_steps, "extract_word_comments"):
        comments = extract_comments(source_copy)
        comments_md.write_text(format_markdown(comments), encoding="utf-8")
        write_json(comments_json, [asdict(comment) for comment in comments])

    with timed_step(profile_steps, "pandoc_docx_to_markdown"):
        run(pandoc_docx_to_markdown(source_copy, markdown, media_dir))
    with timed_step(profile_steps, "seed_revised_markdown"):
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
            "launcher_profile": launcher_profile.name,
        },
    }
    manifest_path = run_dir / "manifest.json"
    with timed_step(profile_steps, "write_agent_scoped_inputs"):
        manifest["agent_inputs"] = write_agent_inputs(run_dir, manifest, markdown, revised_markdown)
    with timed_step(profile_steps, "write_agent_workflow_scaffold"):
        manifest["agent_workflow"] = write_agent_workflow_tasks(run_dir, manifest)
        write_asta_request_template(run_dir, manifest["agent_workflow"])
    with timed_step(profile_steps, "write_manifest"):
        write_json(manifest_path, manifest)
    profile_start = time.perf_counter()
    profile_steps.append({"step": "write_launcher_profile", "seconds": 0.0})
    write_launcher_profile(
        launcher_profile,
        run_dir,
        source,
        manifest,
        profile_steps,
        metadata_audit,
        len(comments),
    )
    profile_steps[-1]["seconds"] = round(time.perf_counter() - profile_start, 4)
    write_launcher_profile(
        launcher_profile,
        run_dir,
        source,
        manifest,
        profile_steps,
        metadata_audit,
        len(comments),
    )

    print(f"Wrote run directory: {run_dir}")
    print(f"Revise markdown: {revised_markdown}")
    print(f"Launcher profile: {launcher_profile}")
    print("Required agent workflow tasks:")
    for task in manifest["agent_workflow"]["task_files"]:
        print(f"  - {run_dir / task}")
    print(f"Agent workflow audit required before finalize: {run_dir / manifest['agent_workflow']['audit_file']}")
    print(f"Finalize with: pandoc-word-revision finalize {manifest_path}")
    return 0


def run_complete(args: argparse.Namespace) -> int:
    source = Path(args.source_docx).resolve()
    output_stem = args.output_stem or source.stem
    run_dir = (Path(args.run_dir) if args.run_dir else source.parent / f"{output_stem}_pandoc_revision_run").resolve()
    manifest_path = run_dir / "manifest.json"

    start_args = argparse.Namespace(
        source_docx=args.source_docx,
        output_stem=args.output_stem,
        run_dir=args.run_dir,
        metadata_ris=args.metadata_ris,
    )
    start(start_args)
    if args.asta_command:
        os.environ["PANDOC_REVISION_ASTA_COMMAND"] = args.asta_command
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    run_asta_preflight_command(run_dir, manifest, args.asta_command)
    run_agent_workflow_command(run_dir, manifest_path, args.agent_command)
    return finalize(argparse.Namespace(manifest=str(manifest_path), asta_command=args.asta_command))


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

    agent_workflow_audit = validate_agent_workflow(run_dir, manifest, revised_markdown)
    asta_metadata_ris = resolve_asta_requests(run_dir, manifest, args.asta_command)

    run(pandoc_markdown_to_docx(revised_markdown, raw_docx, reference_doc))
    stripped_raw = raw_docx.with_name(f"{raw_docx.stem}.stripped{raw_docx.suffix}")
    strip_comments(raw_docx, stripped_raw)
    stripped_raw.replace(raw_docx)

    metadata_name = manifest.get("citation_policy", {}).get("metadata_overlay_ris")
    metadata_ris = asta_metadata_ris or (run_dir / metadata_name if metadata_name else run_dir / "citation_metadata.ris")
    if not metadata_ris.exists():
        metadata_ris = None

    run(reference_list_to_ris_command(raw_docx, ris, metadata_ris, require_metadata_match=metadata_ris is not None))
    check_ris_cmd = reference_list_to_ris_command(raw_docx, ris, metadata_ris, require_metadata_match=metadata_ris is not None)
    check_ris_cmd.append("--check")
    run(check_ris_cmd)
    run([sys.executable, "-m", "citeproc_endnote_uv.docx_plain_numeric_citation_check", str(raw_docx)])
    run(endnote_conversion_command(raw_docx, final_docx, ris))
    run(["unzip", "-t", str(final_docx)])
    run([sys.executable, "-m", "citeproc_endnote_uv.docx_word_sanity", str(final_docx)])
    run([sys.executable, "-m", "citeproc_endnote_uv.docx_endnote_ris_sync", str(final_docx), str(ris)])
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
        "asta_reference_additions": "asta_reference_additions.json" if (run_dir / "asta_reference_additions.json").exists() else None,
        "agent_workflow_audit": manifest["agent_workflow"]["audit_file"],
        "agent_workflow_passes": [item["name"] for item in agent_workflow_audit["passes"]],
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

    run_parser = subparsers.add_parser("run", help="Run the complete start -> agent workflow -> finalize pipeline.")
    run_parser.add_argument("source_docx")
    run_parser.add_argument("--output-stem")
    run_parser.add_argument("--run-dir")
    run_parser.add_argument(
        "--metadata-ris",
        help="Fallback complete RIS metadata overlay, used only when the source DOCX has no embedded EndNote records.",
    )
    run_parser.add_argument(
        "--agent-command",
        help=(
            "Command that runs the required revision agents. If the command string contains placeholders, "
            "{manifest}, {run_dir}, {source_docx}, {source_markdown}, {revised_markdown}, {comments_markdown}, "
            "{comments_json}, {audit_file}, and {asta_requests} are expanded. Otherwise the launcher appends "
            "--manifest and --run-dir. May also be set with PANDOC_REVISION_AGENT_COMMAND."
        ),
    )
    run_parser.add_argument(
        "--asta-command",
        help=(
            "Command used during finalize to resolve pending agent_workflow/asta_requests.json entries. "
            "May also be set with PANDOC_REVISION_ASTA_COMMAND."
        ),
    )
    run_parser.set_defaults(func=run_complete)

    finalize_parser = subparsers.add_parser("finalize", help="Compile revised markdown and finalize DOCX/RIS.")
    finalize_parser.add_argument("manifest")
    finalize_parser.add_argument(
        "--asta-command",
        help=(
            "Command used to resolve pending agent_workflow/asta_requests.json entries. If the command string "
            "contains placeholders, {request_json}, {output_json}, {output_ris}, {run_dir}, and {request_id} "
            "are expanded. Otherwise the launcher appends --request, --output, and --ris arguments. "
            "May also be set with PANDOC_REVISION_ASTA_COMMAND."
        ),
    )
    finalize_parser.set_defaults(func=finalize)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
