#!/usr/bin/env python3
"""Run all required pandoc revision agent passes for a launch manifest.

This runner is intentionally data-local:

- It reads only run artifacts already created by ``pandoc-word-revision start``.
- It invokes a configurable sub-agent command per pass.
- It validates report artifacts and writes ``agent_workflow/agent_workflow_audit.json``.

The launcher remains responsible for passing ``--agent-command``; this module is the
default agent command value that fans out into sub-pass commands.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shlex
import subprocess
from pathlib import Path
from typing import Any

from citeproc_endnote_uv.pandoc_revision_launcher import AGENT_WORKFLOW_PASSES, resolve_asta_requests

MAX_EMBEDDED_CHARS = 50000
DEFAULT_SUBAGENT_COMMAND = (
    'codex exec --model gpt-5.3-codex-spark -c model_reasoning_effort="medium" '
    "--skip-git-repo-check -s workspace-write -C {run_dir} --ephemeral "
    "--output-last-message {report}"
)


def run(command: list[str], *, check: bool = True) -> subprocess.CompletedProcess[str]:
    print("+ " + " ".join(command))
    result = subprocess.run(command, check=check, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    if result.stdout:
        print(result.stdout)
    return result


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def required_passes(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    workflow = manifest.get("agent_workflow", {})
    required = workflow.get("required_passes")
    if isinstance(required, list) and required:
        return [item for item in required if isinstance(item, dict)]
    return [
        {
            "name": item["name"],
            "report": f"agent_workflow/reports/{item['report']}",
            "required_checks": item["required_checks"],
        }
        for item in AGENT_WORKFLOW_PASSES
    ]


def parse_report_checks(report: str, required_checks: list[str]) -> dict[str, bool]:
    checks = {check: False for check in required_checks}
    lines = report.splitlines()

    patterns = [
        re.compile(rf"^\s*(?:[-*]\s*)?`?{re.escape(check)}`?\s*[:=]\s*(true|false|yes|no|1|0)\s*$", re.I)
        for check in required_checks
    ]
    checkbox = re.compile(r"^\s*(?:[-*]\s*)?\[([ xX])\]\s*(.+?)\s*$")

    for line in lines:
        for check, pattern in zip(required_checks, patterns):
            match = pattern.search(line)
            if match:
                checks[check] = match.group(1).lower() in {"true", "yes", "1"}
                break
            match = checkbox.search(line)
            if match and match.group(2).strip() == check:
                checks[check] = match.group(1).lower() in {"x", "X"}
                break

    return checks


def json_payload_candidates(text: str) -> list[str]:
    candidates = re.findall(r"```(?:json)?\s*(\{.*?\})\s*```", text, flags=re.S | re.I)
    stripped = text.strip()
    if stripped.startswith("{") and stripped.endswith("}"):
        candidates.append(stripped)
    marker = "REVISION_PAYLOAD:"
    if marker in text:
        candidates.append(text.split(marker, 1)[1].strip())
    return candidates


def extract_revision_payload(text: str) -> dict[str, Any] | None:
    decoder = json.JSONDecoder()
    for candidate in json_payload_candidates(text):
        try:
            payload = json.loads(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(payload, dict):
            return payload
    for match in re.finditer(r"\{", text):
        tail = text[match.start() :]
        if '"markdown_replacements"' not in tail[:5000]:
            continue
        try:
            payload, _ = decoder.raw_decode(tail)
        except json.JSONDecodeError:
            continue
        if isinstance(payload, dict):
            return payload
    return None


ANNOTATION_REF = '[]{custom-style="annotation reference"}'
CITATION_CLUSTER = re.compile(r"(?<![A-Za-z0-9^])(\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*)")


def normalized_prose(text: str) -> str:
    text = text.replace(ANNOTATION_REF, "")
    text = text.replace("…", "...")
    text = re.sub(r"\^([^^\n]+)\^", r"\1", text)
    text = re.sub(r"^\s{0,3}#{1,6}\s+", "", text)
    return re.sub(r"\s+", " ", text).strip()


def relaxed_prose_key(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", normalized_prose(text).lower()).strip()


def markdownize_plain_citations(text: str) -> str:
    def replacement(match: re.Match[str]) -> str:
        cluster = match.group(1)
        values = [int(value) for value in re.findall(r"\d+", cluster)]
        if "," not in cluster and "-" not in cluster and (not values or values[0] < 10):
            return cluster
        return f"^{cluster}^"

    return CITATION_CLUSTER.sub(replacement, text)


def markdown_paragraphs(text: str) -> list[str]:
    return [paragraph for paragraph in re.split(r"\n\s*\n", text) if paragraph.strip()]


def comment_scope_keys(run_dir: Path, manifest: dict[str, Any]) -> set[str]:
    comments_json = run_dir / manifest["comments"]["json"]
    if not comments_json.exists():
        return set()
    comments = json.loads(comments_json.read_text(encoding="utf-8"))
    keys: set[str] = set()
    for comment in comments if isinstance(comments, list) else []:
        if not isinstance(comment, dict):
            continue
        for field in ("paragraph_text", "anchored_text"):
            value = comment.get(field)
            if isinstance(value, str) and value.strip():
                keys.add(normalized_prose(value))
    return keys


def write_scope_review(run_dir: Path, manifest: dict[str, Any]) -> Path:
    """Write a compact whole-document change-scope artifact for reviewer agents."""
    artifacts = manifest["generated_artifacts"]
    source_markdown = run_dir / artifacts["source_markdown"]
    revised_markdown = run_dir / artifacts["revised_markdown"]
    output = run_dir / "agent_workflow" / "scope_review.md"
    output.parent.mkdir(parents=True, exist_ok=True)

    source_paragraphs = markdown_paragraphs(source_markdown.read_text(encoding="utf-8", errors="ignore"))
    revised_paragraphs = markdown_paragraphs(revised_markdown.read_text(encoding="utf-8", errors="ignore"))
    scoped_keys = comment_scope_keys(run_dir, manifest)
    rows: list[str] = [
        "# Whole-Document Change Scope Review",
        "",
        "This compact artifact lets reviewer passes check whether revisions are limited to Word-commented paragraphs without loading the full manuscript.",
        "",
        f"Source paragraph count: {len(source_paragraphs)}",
        f"Revised paragraph count: {len(revised_paragraphs)}",
        "",
    ]
    changed = 0
    for index, (source, revised) in enumerate(zip(source_paragraphs, revised_paragraphs), start=1):
        if normalized_prose(source) == normalized_prose(revised):
            continue
        changed += 1
        in_comment_scope = normalized_prose(source) in scoped_keys
        rows.extend(
            [
                f"## Changed Paragraph {index}",
                "",
                f"comment_scoped: {'true' if in_comment_scope else 'false'}",
                "",
                "Source:",
                "",
                source.strip(),
                "",
                "Revised:",
                "",
                revised.strip(),
                "",
            ]
        )
    if len(source_paragraphs) != len(revised_paragraphs):
        rows.extend(
            [
                "## Paragraph Count Mismatch",
                "",
                "comment_scoped: false",
                "",
                "The source and revised markdown have different paragraph counts, so reviewer approval requires manual scrutiny.",
                "",
            ]
        )
    if changed == 0 and len(source_paragraphs) == len(revised_paragraphs):
        rows.append("No paragraph-level prose changes detected.")
    output.write_text("\n".join(rows).rstrip() + "\n", encoding="utf-8")
    return output


def replace_normalized_paragraph(text: str, old: str, new: str) -> tuple[str, bool, str | None]:
    paragraphs = re.split(r"(\n\s*\n)", text)
    matches: list[int] = []
    old_norm = normalized_prose(old)
    for index in range(0, len(paragraphs), 2):
        if normalized_prose(paragraphs[index]) == old_norm:
            matches.append(index)
    if len(matches) != 1:
        return text, False, f"normalized paragraph matched {len(matches)} times"
    paragraphs[matches[0]] = markdownize_plain_citations(new)
    return "".join(paragraphs), True, None


def apply_revision_payload(run_dir: Path, manifest: dict[str, Any], report: Path, stdout: str | None) -> dict[str, Any]:
    report_text = report.read_text(encoding="utf-8", errors="ignore") if report.exists() else ""
    payload = extract_revision_payload(report_text)
    if payload is None and stdout:
        payload = extract_revision_payload(stdout)
    if payload is None:
        return {"payload_found": False, "applied": 0, "failed": ["missing revision JSON payload"]}

    replacements = payload.get("markdown_replacements", [])
    if not isinstance(replacements, list):
        return {"payload_found": True, "applied": 0, "failed": ["markdown_replacements is not a list"]}
    if not replacements:
        return {
            "payload_found": True,
            "applied": 0,
            "failed": ["markdown_replacements is empty; implementation pass must revise actionable comments"],
            "deferred": [],
        }

    artifacts = manifest["generated_artifacts"]
    revised_markdown = run_dir / artifacts["revised_markdown"]
    revised_text = revised_markdown.read_text(encoding="utf-8")
    scoped_rel = manifest.get("agent_inputs", {}).get("comment_scoped_revised_markdown")
    scoped_path = run_dir / scoped_rel if scoped_rel else None
    scoped_text = scoped_path.read_text(encoding="utf-8") if scoped_path and scoped_path.exists() else None

    applied = 0
    failed: list[str] = []
    deferred: list[str] = []
    for index, item in enumerate(replacements, start=1):
        if not isinstance(item, dict):
            failed.append(f"replacement {index} is not an object")
            continue
        old = item.get("old")
        new = item.get("new")
        if not isinstance(old, str) or not old:
            failed.append(f"replacement {index} has no exact old string")
            continue
        if not isinstance(new, str):
            failed.append(f"replacement {index} has no new string")
            continue
        if old not in revised_text and (
            normalized_prose(new) in normalized_prose(revised_text)
            or relaxed_prose_key(new) in relaxed_prose_key(revised_text)
        ):
            applied += 1
            continue
        count = revised_text.count(old)
        if count == 1:
            revised_text = revised_text.replace(old, new, 1)
            applied += 1
        else:
            revised_text, replaced, reason = replace_normalized_paragraph(revised_text, old, new)
            if not replaced:
                if replacement_is_deferred_to_asta(old, new, payload):
                    deferred.append(f"replacement {index} deferred to Asta evidence resolution; {reason}")
                    continue
                failed.append(f"replacement {index} matched {count} times in revised markdown; {reason}")
                continue
            applied += 1
        if scoped_text is not None and old in scoped_text:
            scoped_text = scoped_text.replace(old, new, 1)
        elif scoped_text is not None:
            scoped_text, _, _ = replace_normalized_paragraph(scoped_text, old, new)

    if applied:
        revised_markdown.write_text(revised_text, encoding="utf-8")
        if scoped_path and scoped_text is not None:
            scoped_path.write_text(scoped_text, encoding="utf-8")

    requests = payload.get("asta_requests", [])
    if isinstance(requests, list) and requests:
        merge_asta_requests(run_dir, manifest, requests)

    return {"payload_found": True, "applied": applied, "failed": failed, "deferred": deferred}


def replacement_is_deferred_to_asta(old: str, new: str, payload: dict[str, Any]) -> bool:
    requests = payload.get("asta_requests", [])
    if not isinstance(requests, list) or not requests:
        return False
    text = f"{old}\n{new}".lower()
    if not any(marker in text for marker in ["requires citation", "should cite", "requires asta", "h3k27ac"]):
        return False
    request_text = " ".join(
        " ".join(str(value) for value in item.values()).lower()
        for item in requests
        if isinstance(item, dict)
    )
    return "h3k27ac" in text and "h3k27ac" in request_text


def merge_asta_requests(run_dir: Path, manifest: dict[str, Any], requests: list[Any]) -> None:
    workflow = manifest["agent_workflow"]
    requests_path = run_dir / workflow["asta_requests"]
    if requests_path.exists():
        ledger = json.loads(requests_path.read_text(encoding="utf-8"))
    else:
        ledger = {"version": 1, "requests": []}
    existing = ledger.get("requests")
    if not isinstance(existing, list):
        existing = []
        ledger["requests"] = existing
    existing_ids = {item.get("id") for item in existing if isinstance(item, dict)}
    existing_keys = {
        (item.get("comment_id"), item.get("request") or item.get("query") or item.get("claim"))
        for item in existing
        if isinstance(item, dict)
    }
    for index, item in enumerate(requests, start=1):
        if not isinstance(item, dict):
            continue
        request = dict(item)
        request_key = (request.get("comment_id"), request.get("request") or request.get("query") or request.get("claim"))
        if request_key in existing_keys:
            continue
        request.setdefault("required", True)
        request.setdefault("status", "pending")
        request.setdefault("id", f"revision-request-{len(existing) + index}")
        if request["id"] in existing_ids:
            continue
        existing.append(request)
        existing_ids.add(request["id"])
        existing_keys.add(request_key)
    requests_path.parent.mkdir(parents=True, exist_ok=True)
    requests_path.write_text(json.dumps(ledger, indent=2) + "\n", encoding="utf-8")


def pending_required_asta_requests(run_dir: Path, manifest: dict[str, Any]) -> list[dict[str, Any]]:
    workflow = manifest.get("agent_workflow", {})
    requests_rel = workflow.get("asta_requests", "agent_workflow/asta_requests.json")
    requests_path = run_dir / str(requests_rel)
    if not requests_path.exists():
        return []
    ledger = json.loads(requests_path.read_text(encoding="utf-8"))
    requests = ledger.get("requests", [])
    if not isinstance(requests, list):
        return []
    return [
        item
        for item in requests
        if isinstance(item, dict)
        and item.get("required", True)
        and item.get("status", "pending") not in {"resolved", "not_needed"}
    ]


def asta_resolver_configured() -> bool:
    return bool(os.environ.get("PANDOC_REVISION_ASTA_COMMAND"))


def resolve_pending_asta_requests(run_dir: Path, manifest: dict[str, Any], *, reason: str) -> bool:
    pending = pending_required_asta_requests(run_dir, manifest)
    if not pending:
        return False
    if not asta_resolver_configured():
        ids = ", ".join(str(item.get("id", f"request-{index}")) for index, item in enumerate(pending, start=1))
        raise SystemExit(
            f"{reason} created required Asta requests, but no Asta resolver is configured. "
            "Set PANDOC_REVISION_ASTA_COMMAND or pass --asta-command. Pending request ids: " + ids
        )
    resolve_asta_requests(run_dir, manifest, os.environ.get("PANDOC_REVISION_ASTA_COMMAND"))
    return True


def ensure_exists(path: Path) -> None:
    if not path.exists():
        raise SystemExit(f"Missing required file: {path}")


def read_embedded(path: Path, label: str, max_chars: int = MAX_EMBEDDED_CHARS) -> str:
    text = path.read_text(encoding="utf-8", errors="ignore")
    if len(text) <= max_chars:
        return f"## {label}\nPath: `{path}`\n\n```text\n{text}\n```\n"
    head = text[: max_chars // 2]
    tail = text[-max_chars // 2 :]
    return (
        f"## {label}\nPath: `{path}`\n"
        f"Embedded excerpt is truncated to {max_chars} characters from {len(text)} total characters.\n\n"
        f"```text\n{head}\n\n[... truncated ...]\n\n{tail}\n```\n"
    )


def embedded_inputs_for_pass(run_dir: Path, manifest: dict[str, Any], name: str) -> str:
    inputs = manifest.get("agent_inputs", {})
    policy = inputs.get("pass_input_policy", {}).get(name, {})
    recommended = policy.get("recommended_inputs")
    if not isinstance(recommended, list) or not recommended:
        recommended = [
            manifest["comments"]["markdown"],
            inputs.get("comment_scoped_source_markdown"),
            inputs.get("comment_scoped_revised_markdown"),
            "manifest.json",
        ]
    chunks: list[str] = []
    seen: set[Path] = set()
    for rel in recommended:
        if not rel:
            continue
        path = (run_dir / str(rel)).resolve()
        try:
            path.relative_to(run_dir.resolve())
        except ValueError:
            continue
        if path in seen or not path.exists() or not path.is_file():
            continue
        seen.add(path)
        chunks.append(read_embedded(path, str(rel)))
    workflow = manifest.get("agent_workflow", {})
    extra_paths: list[tuple[Path, str]] = []
    requests_rel = workflow.get("asta_requests")
    if requests_rel:
        extra_paths.append((run_dir / str(requests_rel), str(requests_rel)))
    asta_dir = run_dir / str(workflow.get("asta_resolutions_dir", "agent_workflow/asta"))
    if asta_dir.exists():
        for path in sorted(asta_dir.glob("responses/*.json")):
            if path.name.endswith(".asta.json"):
                continue
            extra_paths.append((path, str(path.relative_to(run_dir))))
        additions = run_dir / "asta_reference_additions.json"
        if additions.exists():
            extra_paths.append((additions, additions.name))
    if name == "rigor_critique":
        scope_review = run_dir / "agent_workflow" / "scope_review.md"
        if scope_review.exists():
            extra_paths.append((scope_review, str(scope_review.relative_to(run_dir))))
    for path, label in extra_paths:
        path = path.resolve()
        try:
            path.relative_to(run_dir.resolve())
        except ValueError:
            continue
        if path in seen or not path.exists() or not path.is_file():
            continue
        seen.add(path)
        chunks.append(read_embedded(path, label, max_chars=20000))
    return "\n".join(chunks)


def pass_prompt(manifest_path: Path, run_dir: Path, manifest: dict[str, Any], name: str, task: Path, report: Path) -> str:
    workflow = manifest["agent_workflow"]
    artifacts = manifest["generated_artifacts"]
    comments = manifest["comments"]
    source_docx = run_dir / manifest["source_docx"]
    source_markdown = run_dir / artifacts["source_markdown"]
    revised_markdown = run_dir / artifacts["revised_markdown"]
    comments_markdown = run_dir / comments["markdown"]
    comments_json = run_dir / comments["json"]
    asta_requests = run_dir / workflow["asta_requests"]

    task_note = task.read_text(encoding="utf-8", errors="ignore").strip()
    embedded_inputs = embedded_inputs_for_pass(run_dir, manifest, name)
    required = next(
        (item["required_checks"] for item in required_passes(manifest) if item["name"] == name),
        [],
    )
    required_skills = next(
        (item.get("required_skills", []) for item in required_passes(manifest) if item["name"] == name),
        [],
    )
    skill_instruction = ""
    if required_skills:
        skill_names = ", ".join(f"`{skill}`" for skill in required_skills)
        skill_checks = ", ".join(f"`{skill.replace('-', '_')}_skill_used: true`" for skill in required_skills)
        skill_instruction = (
            "Required Codex skills: "
            f"{skill_names}.\n"
            "Use these skills the same way the main Codex agent would: invoke each named skill before doing this pass, "
            "follow its SKILL.md guidance, and make the skill use explicit in the final report. If a required skill "
            "is unavailable in the nested session, mark its check false and explain why. Required skill checks: "
            f"{skill_checks}.\n\n"
        )

    return (
        "You are executing one pass of an agent workflow for a manuscript revision.\n"
        f"Pass: {name}\n"
        f"Manifest: {manifest_path}\n"
        f"Run dir: {run_dir}\n"
        f"Source DOCX: {source_docx}\n"
        f"Source markdown: {source_markdown}\n"
        f"Revised markdown: {revised_markdown}\n"
        f"Comments markdown: {comments_markdown}\n"
        f"Comments JSON: {comments_json}\n"
        f"Asta requests: {asta_requests}\n"
        f"Report target: {report}\n"
        f"Required checks: {', '.join(required)}\n\n"
        "Use the task file context below and return a concise report as your final answer.\n"
        "Do not call shell commands. Do not use file-editing tools. Do not attempt to write files. "
        "The parent runner captures your final answer into the report target with --output-last-message, "
        "then applies any structured payload inside the run directory.\n\n"
        f"{skill_instruction}"
        f"{task_note}\n\n"
        "Nested shell commands and direct file writes may be unavailable in sandboxed subagents. The run-local "
        "inputs needed for this pass are embedded below. Use these embedded inputs only.\n\n"
        f"{embedded_inputs}\n"
        "If this pass identifies a fixable wording or citation-support problem that should block approval, "
        "include a fenced JSON object with exact markdown replacements so the parent runner can apply the fix "
        "and rerun the pass. For the `revision_implementation` pass, this JSON payload is required.\n"
        "```json\n"
        "{\"markdown_replacements\":[{\"old\":\"exact original markdown\",\"new\":\"replacement markdown\"}],"
        "\"asta_requests\":[]}\n"
        "```\n"
        "The parent runner applies these exact replacements inside the run directory after the subagent exits.\n\n"
        "Report must be non-empty and include each required check as either:\n"
        "- check_name: true/false\n"
        "- [x] check_name\n\n"
        "Return only the final report content. The final report must contain the required checks and any JSON payload."
    )


def resolve_subagent_command(template: str, context: dict[str, str], prompt: str) -> list[str]:
    if "{" in template and "}" in template:
        command = template.format(**context, prompt=prompt)
        parts = shlex.split(command)
        if "{prompt}" not in template:
            parts.append(prompt)
    else:
        parts = shlex.split(template) + [prompt]
    return parts


def run_single_pass(
    run_dir: Path,
    manifest_path: Path,
    manifest: dict[str, Any],
    pass_definition: dict[str, Any],
    subagent_template: str,
) -> dict[str, Any]:
    name = pass_definition["name"]
    report_rel = pass_definition["report"]
    task = run_dir / "agent_workflow" / "tasks" / f"{name}.md"
    report = run_dir / report_rel
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(
        "\n".join(
            [
                f"# {name}",
                "",
                "status: pending",
                *(f"{check}: false" for check in pass_definition.get("required_checks", [])),
                "",
            ]
        ),
        encoding="utf-8",
    )

    required = list(pass_definition.get("required_checks", []))
    if not task.exists():
        raise SystemExit(f"Missing pass task file: {task}")

    write_scope_review(run_dir, manifest)
    prompt = pass_prompt(manifest_path, run_dir, manifest, name, task, report)
    context = {
        "manifest": str(manifest_path),
        "run_dir": str(run_dir),
        "source_docx": str(run_dir / manifest["source_docx"]),
        "source_markdown": str(run_dir / manifest["generated_artifacts"]["source_markdown"]),
        "revised_markdown": str(run_dir / manifest["generated_artifacts"]["revised_markdown"]),
        "comments_markdown": str(run_dir / manifest["comments"]["markdown"]),
        "comments_json": str(run_dir / manifest["comments"]["json"]),
        "audit_file": str(run_dir / manifest["agent_workflow"]["audit_file"]),
        "asta_requests": str(run_dir / manifest["agent_workflow"]["asta_requests"]),
        "task": str(task),
        "report": str(report),
        "pass_name": name,
    }

    result = run(resolve_subagent_command(subagent_template, context, prompt))
    if not report.exists() or not report.read_text(encoding="utf-8").strip():
        raise SystemExit(f"Pass {name} failed to write report: {report}")

    report_text = report.read_text(encoding="utf-8")
    checks = parse_report_checks(report_text, required)
    if not all(checks.values()) and result.stdout:
        stdout_checks = parse_report_checks(result.stdout, required)
        if all(stdout_checks.values()):
            report.write_text(result.stdout, encoding="utf-8")
            report_text = result.stdout
            checks = stdout_checks
    application: dict[str, Any] | None = None
    payload = extract_revision_payload(report_text)
    if name == "revision_implementation" or (payload and payload.get("markdown_replacements")):
        application = apply_revision_payload(run_dir, manifest, report, result.stdout)
        if "revisions_applied" in checks:
            checks["revisions_applied"] = checks["revisions_applied"] and bool(application["applied"]) and not application["failed"]
    return {
        "name": name,
        "status": "completed" if all(checks.values()) else "blocked",
        "report": report_rel,
        "checks": checks,
        **({"application": application} if application is not None else {}),
    }


def write_audit(
    run_dir: Path,
    manifest: dict[str, Any],
    passes: list[dict[str, Any]],
) -> Path:
    workflow = manifest["agent_workflow"]
    artifacts = manifest["generated_artifacts"]
    revised_markdown = run_dir / artifacts["revised_markdown"]

    overall = {
        "all_comments_addressed": False,
        "modified_claims_have_adjacent_citation_or_resolution": False,
        "uncommented_changes_justified": False,
        "citation_integrity_reviewed": False,
        "ready_for_finalize": False,
    }
    if passes:
        pass_status = {item["name"]: item for item in passes}
        check_implementation = pass_status.get("revision_implementation", {})
        check_plan = pass_status.get("comment_interpretation_and_revision_planning", {})
        check_evidence = pass_status.get("evidence_and_specificity", {})
        check_rigor = pass_status.get("rigor_critique", {})
        overall["all_comments_addressed"] = (
            check_implementation.get("checks", {}).get("revisions_applied", False)
            and check_plan.get("checks", {}).get("comments_addressed", False)
        )
        overall["modified_claims_have_adjacent_citation_or_resolution"] = (
            check_evidence.get("checks", {}).get("modified_claims_citation_checked", False)
            and check_evidence.get("checks", {}).get("unsupported_claims_resolved", False)
        )
        overall["uncommented_changes_justified"] = (
            check_rigor.get("checks", {}).get("rigor_approved", False)
            and check_rigor.get("checks", {}).get("uncommented_changes_reviewed", False)
        )
        overall["citation_integrity_reviewed"] = (
            check_evidence.get("checks", {}).get("source_docx_only", False)
            if check_evidence
            else False
        )
        overall["ready_for_finalize"] = all(item["status"] == "completed" for item in passes)

    audit = {
        "workflow": "pandoc-word-revision-agent-workflow",
        "source_sha256": manifest["source_sha256"],
        "revised_markdown": artifacts["revised_markdown"],
        "revised_markdown_sha256": sha256(revised_markdown),
        "passes": passes,
        "overall": overall,
        "source_path": manifest["source_docx"],
    }
    audit_path = run_dir / workflow["audit_file"]
    audit_path.write_text(json.dumps(audit, indent=2) + "\n", encoding="utf-8")
    return audit_path


def run_all(manifest: Path, run_dir: Path, subagent_command: str | None) -> int:
    manifest_data = json.loads(manifest.read_text(encoding="utf-8"))
    workflow = manifest_data.get("agent_workflow")
    if not workflow:
        raise SystemExit("Manifest is missing agent_workflow data.")

    source_docx = run_dir / manifest_data["source_docx"]
    artifacts = manifest_data["generated_artifacts"]
    ensure_exists(source_docx)
    source_md = run_dir / artifacts["source_markdown"]
    revised_md = run_dir / artifacts["revised_markdown"]
    comments_md = run_dir / manifest_data["comments"]["markdown"]
    comments_json = run_dir / manifest_data["comments"]["json"]
    ensure_exists(source_md)
    ensure_exists(revised_md)
    ensure_exists(comments_md)
    ensure_exists(comments_json)

    passes = required_passes(manifest_data)
    if not passes:
        raise SystemExit("No passes are configured in the manifest.")
    if not subagent_command:
        subagent_command = os.environ.get("PANDOC_REVISION_SUBAGENT_COMMAND")
    if not subagent_command:
        subagent_command = DEFAULT_SUBAGENT_COMMAND

    completed: list[dict[str, Any]] = []
    for pass_definition in passes:
        result = run_single_pass(run_dir, manifest, manifest_data, pass_definition, subagent_command)
        application = result.get("application") if isinstance(result.get("application"), dict) else None
        if (
            pass_definition["name"] != "revision_implementation"
            and application
            and application.get("applied")
            and not application.get("failed")
        ):
            result = run_single_pass(run_dir, manifest, manifest_data, pass_definition, subagent_command)
        completed.append(result)
        if pass_definition["name"] == "revision_implementation" and result["status"] != "completed":
            write_audit(run_dir, manifest_data, completed)
            raise SystemExit(
                "Revision implementation pass did not complete; stopping before reviewer passes. "
                f"See {run_dir / result['report']}."
            )
        if pass_definition["name"] == "revision_implementation":
            if resolve_pending_asta_requests(run_dir, manifest_data, reason="Revision implementation"):
                result = run_single_pass(run_dir, manifest, manifest_data, pass_definition, subagent_command)
                completed[-1] = result
                if result["status"] != "completed":
                    write_audit(run_dir, manifest_data, completed)
                    raise SystemExit(
                        "Revision implementation remained blocked after Asta resolution; stopping before reviewer passes. "
                        f"See {run_dir / result['report']}."
                    )
        if pass_definition["name"] == "evidence_and_specificity":
            if resolve_pending_asta_requests(run_dir, manifest_data, reason="Evidence review"):
                result = run_single_pass(run_dir, manifest, manifest_data, pass_definition, subagent_command)
                completed[-1] = result

    write_audit(run_dir, manifest_data, completed)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True)
    parser.add_argument("--run-dir", required=True)
    parser.add_argument("--subagent-command", help="Command template used for each workflow pass.")
    parser.add_argument(
        "--asta-command",
        help="Compatibility passthrough; not consumed by this runner except being retained in environment.",
    )
    args = parser.parse_args(argv)

    if args.asta_command:
        os.environ.setdefault("PANDOC_REVISION_ASTA_COMMAND", args.asta_command)

    manifest = Path(args.manifest).resolve()
    run_dir = Path(args.run_dir).resolve()
    if not manifest.exists():
        raise SystemExit(f"Manifest not found: {manifest}")
    if not run_dir.exists():
        raise SystemExit(f"Run dir not found: {run_dir}")

    return run_all(manifest, run_dir, args.subagent_command)


if __name__ == "__main__":
    raise SystemExit(main())
