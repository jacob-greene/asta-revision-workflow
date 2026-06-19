from pathlib import Path

import json

from citeproc_endnote_uv.pandoc_revision_agent_runner import run_all


def test_agent_runner_executes_all_passes_and_writes_audit(tmp_path, monkeypatch):
    manifest = tmp_path / "manifest.json"
    run_dir = tmp_path / "run"
    run_dir.mkdir()

    source_docx = run_dir / "source.docx"
    source_markdown = run_dir / "source.md"
    revised_markdown = run_dir / "revised.md"
    comments_md = run_dir / "source.comments.md"
    comments_json = run_dir / "source.comments.json"
    source_docx.write_text("", encoding="utf-8")
    source_markdown.write_text("# source\n", encoding="utf-8")
    revised_markdown.write_text("# revised\n", encoding="utf-8")
    comments_md.write_text("# comments\n", encoding="utf-8")
    comments_json.write_text("[]", encoding="utf-8")

    manifest.write_text(
        json.dumps(
            {
                "workflow": "pandoc-word-revision",
                "source_docx": source_docx.name,
                "source_sha256": "abc",
                "generated_artifacts": {
                    "source_markdown": source_markdown.name,
                    "revised_markdown": revised_markdown.name,
                    "media_dir": "media",
                },
                "comments": {"markdown": comments_md.name, "json": comments_json.name},
                "agent_workflow": {
                    "audit_file": "agent_workflow/agent_workflow_audit.json",
                    "asta_requests": "agent_workflow/asta_requests.json",
                    "required_passes": [
                        {
                            "name": "revision_implementation",
                            "report": "agent_workflow/reports/revision_implementation_report.md",
                            "required_checks": ["revisions_applied", "comment_scope_preserved", "source_docx_only"],
                        },
                        {
                            "name": "comment_interpretation_and_revision_planning",
                            "report": "agent_workflow/reports/comment_plan_report.md",
                            "required_checks": ["comments_addressed", "revision_scope_defined", "source_docx_only"],
                        },
                        {
                            "name": "evidence_and_specificity",
                            "report": "agent_workflow/reports/evidence_specificity_report.md",
                            "required_checks": [
                                "modified_claims_citation_checked",
                                "unsupported_claims_resolved",
                                "source_docx_only",
                            ],
                        },
                        {
                            "name": "rigor_critique",
                            "report": "agent_workflow/reports/rigor_critique_report.md",
                            "required_checks": [
                                "rigor_approved",
                                "new_knowledge_claims_skeptically_reviewed",
                                "uncommented_changes_reviewed",
                            ],
                        },
                        {
                            "name": "tone_and_concision",
                            "report": "agent_workflow/reports/tone_concision_report.md",
                            "required_checks": ["tone_reviewed", "redundancy_checked", "comment_scope_preserved"],
                        },
                    ],
                },
            },
            indent=2,
        )
    )

    (run_dir / "agent_workflow" / "tasks").mkdir(parents=True)
    for name in [
        "revision_implementation",
        "comment_interpretation_and_revision_planning",
        "evidence_and_specificity",
        "rigor_critique",
        "tone_and_concision",
    ]:
        (run_dir / "agent_workflow" / "tasks" / f"{name}.md").write_text("# task\n", encoding="utf-8")
    (run_dir / "agent_workflow" / "asta_requests.json").write_text('{"version":1,"requests":[]}', encoding="utf-8")

    calls: list[list[str]] = []

    def fake_run(command):
        calls.append(command)
        # infer pass from command placeholders (last positional token is report path for fallback format)
        for token in command:
            if token.endswith("revision_implementation_report.md"):
                pass_name = "revision_implementation"
                report = Path(token)
                break
            if token.endswith("comment_plan_report.md"):
                pass_name = "comment_interpretation_and_revision_planning"
                report = Path(token)
                break
            if token.endswith("evidence_specificity_report.md"):
                pass_name = "evidence_and_specificity"
                report = Path(token)
                break
            if token.endswith("rigor_critique_report.md"):
                pass_name = "rigor_critique"
                report = Path(token)
                break
            if token.endswith("tone_concision_report.md"):
                pass_name = "tone_and_concision"
                report = Path(token)
                break
        else:
            raise AssertionError("missing report path in fake command")

        report_lines = {
            "revision_implementation": [
                "revisions_applied: true",
                "comment_scope_preserved: true",
                "source_docx_only: true",
                '```json\n{"markdown_replacements":[{"old":"# revised\\n","new":"# revised\\n\\nImplemented.\\n"}],"asta_requests":[]}\n```',
            ],
            "comment_interpretation_and_revision_planning": [
                "comments_addressed: true",
                "revision_scope_defined: true",
                "source_docx_only: true",
            ],
            "evidence_and_specificity": [
                "modified_claims_citation_checked: true",
                "unsupported_claims_resolved: true",
                "source_docx_only: true",
            ],
            "rigor_critique": [
                "rigor_approved: true",
                "new_knowledge_claims_skeptically_reviewed: true",
                "uncommented_changes_reviewed: true",
            ],
            "tone_and_concision": [
                "tone_reviewed: true",
                "redundancy_checked: true",
                "comment_scope_preserved: true",
            ],
        }[pass_name]
        report.write_text("\n".join(report_lines), encoding="utf-8")
        return type("Result", (), {"stdout": ""})()

    def mock_run(command, check=True):
        return fake_run(command)

    monkeypatch.setattr("citeproc_endnote_uv.pandoc_revision_agent_runner.run", mock_run)

    assert run_all(manifest, run_dir, "agent {run_dir} --pass-name {pass_name} --report {report}") == 0
    assert len(calls) == 5
    assert "Implemented." in revised_markdown.read_text(encoding="utf-8")
    audit = json.loads((run_dir / "agent_workflow" / "agent_workflow_audit.json").read_text(encoding="utf-8"))
    assert audit["overall"]["ready_for_finalize"]
    assert all(item["status"] == "completed" for item in audit["passes"])
