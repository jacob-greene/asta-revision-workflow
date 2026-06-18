from pathlib import Path
import json

import pytest

from citeproc_endnote_uv.pandoc_revision_launcher import (
    AGENT_WORKFLOW_PASSES,
    PANDOC_FROM,
    PANDOC_TO,
    ensure_inside_run_dir,
    pandoc_docx_to_markdown,
    pandoc_markdown_to_docx,
    sha256,
    validate_agent_workflow,
    write_agent_inputs,
    write_agent_workflow_tasks,
    write_json,
    write_launcher_profile,
)


def test_pandoc_docx_to_markdown_preserves_styles_and_extracts_media(tmp_path):
    source = tmp_path / "source.docx"
    markdown = tmp_path / "source.md"
    media = tmp_path / "media"

    command = pandoc_docx_to_markdown(source, markdown, media)

    assert command == [
        "pandoc",
        "-f",
        PANDOC_FROM,
        "-t",
        PANDOC_TO,
        "--wrap=none",
        f"--extract-media={media}",
        str(source),
        "-o",
        str(markdown),
    ]


def test_pandoc_markdown_to_docx_uses_saved_reference_doc(tmp_path):
    markdown = tmp_path / "revised.md"
    output = tmp_path / "output.docx"
    reference = tmp_path / "style-reference.docx"

    command = pandoc_markdown_to_docx(markdown, output, reference)

    assert command == [
        "pandoc",
        "-f",
        PANDOC_TO,
        "-t",
        "docx",
        "--wrap=none",
        f"--reference-doc={reference}",
        str(markdown),
        "-o",
        str(output),
    ]


def test_finalize_inputs_must_remain_inside_run_dir(tmp_path):
    run_dir = tmp_path / "run"
    run_dir.mkdir()
    inside = run_dir / "revised.md"
    outside = tmp_path / "outside.md"

    assert ensure_inside_run_dir(inside, run_dir, "revised markdown") == inside.resolve()
    with pytest.raises(SystemExit, match="must be inside the run directory"):
        ensure_inside_run_dir(outside, run_dir, "revised markdown")


def agent_manifest(tmp_path):
    revised = tmp_path / "draft.revised.md"
    revised.write_text("Revised text.\n", encoding="utf-8")
    return revised, {
        "workflow": "pandoc-word-revision",
        "source_docx": "source.docx",
        "source_sha256": "source-hash",
        "pandoc": {"reference_doc": "style-reference.docx"},
        "comments": {"markdown": "draft.comments.md", "json": "draft.comments.json", "count": 1},
        "citation_policy": {"metadata_overlay_ris": "citation_metadata.ris", "metadata_audit": "citation_metadata_audit.json"},
        "generated_artifacts": {
            "source_markdown": "draft.source.md",
            "revised_markdown": revised.name,
            "media_dir": "media",
            "raw_docx": "draft.raw.docx",
            "final_docx": "draft.docx",
            "ris": "draft.ris",
        },
    }


def write_agent_fixture_files(tmp_path, manifest):
    (tmp_path / manifest["comments"]["markdown"]).write_text("# Comments\n", encoding="utf-8")
    write_json(
        tmp_path / manifest["comments"]["json"],
        [
            {
                "comment_id": "7",
                "comment_text": "Clarify this claim.",
                "paragraph_index": 3,
                "paragraph_text": "This source paragraph needs a more precise claim.",
                "anchored_text": "more precise claim",
            }
        ],
    )
    (tmp_path / manifest["citation_policy"]["metadata_overlay_ris"]).write_text("TY  - JOUR\nER  -\n", encoding="utf-8")
    (tmp_path / manifest["generated_artifacts"]["source_markdown"]).write_text("Full source markdown.\n", encoding="utf-8")


def completed_agent_audit(manifest, revised):
    return {
        "workflow": "pandoc-word-revision-agent-workflow",
        "source_sha256": manifest["source_sha256"],
        "revised_markdown": manifest["generated_artifacts"]["revised_markdown"],
        "revised_markdown_sha256": sha256(revised),
        "passes": [
            {
                "name": workflow_pass["name"],
                "status": "completed",
                "report": f"agent_workflow/reports/{workflow_pass['report']}",
                "checks": {check: True for check in workflow_pass["required_checks"]},
            }
            for workflow_pass in AGENT_WORKFLOW_PASSES
        ],
        "overall": {
            "all_comments_addressed": True,
            "modified_claims_have_adjacent_citation_or_resolution": True,
            "uncommented_changes_justified": True,
            "citation_integrity_reviewed": True,
            "ready_for_finalize": True,
        },
    }


def test_agent_workflow_scaffold_creates_tasks_and_template(tmp_path):
    revised, manifest = agent_manifest(tmp_path)
    write_agent_fixture_files(tmp_path, manifest)
    manifest["agent_inputs"] = write_agent_inputs(
        tmp_path, manifest, tmp_path / manifest["generated_artifacts"]["source_markdown"], revised
    )

    workflow = write_agent_workflow_tasks(tmp_path, manifest)

    assert workflow["required"] is True
    assert len(workflow["task_files"]) == 4
    assert len(workflow["required_reports"]) == 4
    assert (tmp_path / workflow["audit_template"]).exists()
    first_task = (tmp_path / workflow["task_files"][0]).read_text(encoding="utf-8")
    assert "Required Checks" in first_task
    assert "draft.revised.md" in first_task
    assert "Recommended Minimal Inputs" in first_task
    assert "citation_metadata.ris" in first_task


def test_agent_inputs_use_comment_scope_and_avoid_ris_for_tone(tmp_path):
    revised, manifest = agent_manifest(tmp_path)
    write_agent_fixture_files(tmp_path, manifest)

    agent_inputs = write_agent_inputs(tmp_path, manifest, tmp_path / manifest["generated_artifacts"]["source_markdown"], revised)

    scoped = (tmp_path / agent_inputs["comment_scoped_source_markdown"]).read_text(encoding="utf-8")
    assert "Clarify this claim." in scoped
    assert "This source paragraph needs" in scoped
    tone_policy = agent_inputs["pass_input_policy"]["tone_and_concision"]
    assert "citation_metadata.ris" in tone_policy["avoid_by_default"]


def test_launcher_profile_estimates_scoped_token_savings(tmp_path):
    revised, manifest = agent_manifest(tmp_path)
    write_agent_fixture_files(tmp_path, manifest)
    manifest["agent_inputs"] = write_agent_inputs(
        tmp_path, manifest, tmp_path / manifest["generated_artifacts"]["source_markdown"], revised
    )
    manifest["agent_workflow"] = write_agent_workflow_tasks(tmp_path, manifest)
    write_json(tmp_path / "manifest.json", manifest)
    profile_path = tmp_path / "launcher_profile.json"

    write_launcher_profile(
        profile_path,
        tmp_path,
        tmp_path / "source.docx",
        manifest,
        [{"step": "example", "seconds": 0.1}],
        {"embedded_record_count": 1},
        1,
    )

    profile = json.loads(profile_path.read_text(encoding="utf-8"))
    assert profile["agent_pass_estimates"]
    assert profile["four_pass_estimated_token_savings"] >= 0


def test_agent_workflow_validation_requires_audit(tmp_path):
    revised, manifest = agent_manifest(tmp_path)
    manifest["agent_workflow"] = write_agent_workflow_tasks(tmp_path, manifest)

    with pytest.raises(SystemExit, match="Missing required agent workflow audit"):
        validate_agent_workflow(tmp_path, manifest, revised)


def test_agent_workflow_validation_accepts_complete_audit(tmp_path):
    revised, manifest = agent_manifest(tmp_path)
    manifest["agent_workflow"] = write_agent_workflow_tasks(tmp_path, manifest)
    for report in manifest["agent_workflow"]["required_reports"]:
        report_path = tmp_path / report
        report_path.write_text("completed\n", encoding="utf-8")
    audit_path = tmp_path / manifest["agent_workflow"]["audit_file"]
    audit = completed_agent_audit(manifest, revised)
    audit_path.write_text(json.dumps(audit), encoding="utf-8")

    assert validate_agent_workflow(tmp_path, manifest, revised) == audit


def test_agent_workflow_validation_hashes_revised_markdown(tmp_path):
    revised, manifest = agent_manifest(tmp_path)
    manifest["agent_workflow"] = write_agent_workflow_tasks(tmp_path, manifest)
    for report in manifest["agent_workflow"]["required_reports"]:
        (tmp_path / report).write_text("completed\n", encoding="utf-8")
    audit = completed_agent_audit(manifest, revised)
    revised.write_text("Changed after review.\n", encoding="utf-8")
    (tmp_path / manifest["agent_workflow"]["audit_file"]).write_text(json.dumps(audit), encoding="utf-8")

    with pytest.raises(SystemExit, match="hash does not match"):
        validate_agent_workflow(tmp_path, manifest, revised)
