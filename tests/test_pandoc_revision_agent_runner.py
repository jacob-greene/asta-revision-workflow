from pathlib import Path
import json

import pytest

from asta_revision_workflow.pandoc_revision_agent_runner import run_all


def test_agent_runner_executes_four_passes_and_writes_audit(tmp_path, monkeypatch):
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
                            "required_checks": [
                                "revisions_applied",
                                "comment_scope_preserved",
                                "source_docx_only",
                                "draft_scientific_paper_skill_used",
                                "edit_scientific_prose_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper", "edit-scientific-prose"],
                        },
                        {
                            "name": "asta_query_and_collation",
                            "report": "agent_workflow/reports/asta_query_and_collation_report.md",
                            "required_checks": [
                                "modified_claims_citation_checked",
                                "unsupported_claims_resolved",
                                "asta_requests_collated",
                                "source_docx_only",
                                "draft_scientific_paper_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper"],
                        },
                        {
                            "name": "rigor_critique",
                            "report": "agent_workflow/reports/rigor_critique_report.md",
                            "required_checks": [
                                "rigor_approved",
                                "new_knowledge_claims_skeptically_reviewed",
                                "uncommented_changes_reviewed",
                                "draft_scientific_paper_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper"],
                        },
                        {
                            "name": "tone_and_concision",
                            "report": "agent_workflow/reports/tone_concision_report.md",
                            "required_checks": [
                                "tone_reviewed",
                                "redundancy_checked",
                                "comment_scope_preserved",
                                "edit_scientific_prose_skill_used",
                            ],
                            "required_skills": ["edit-scientific-prose"],
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
        "asta_query_and_collation",
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
            if token.endswith("asta_query_and_collation_report.md"):
                pass_name = "asta_query_and_collation"
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
                    "draft_scientific_paper_skill_used: true",
                    "edit_scientific_prose_skill_used: true",
                    '```json\n{"markdown_replacements":[{"old":"# revised\\n","new":"# revised\\n\\nImplemented.\\n"}],"asta_requests":[]}\n```',
                ],
                "asta_query_and_collation": [
                    "modified_claims_citation_checked: true",
                    "unsupported_claims_resolved: true",
                    "asta_requests_collated: true",
                    "source_docx_only: true",
                    "draft_scientific_paper_skill_used: true",
                ],
                "rigor_critique": [
                    "rigor_approved: true",
                    "new_knowledge_claims_skeptically_reviewed: true",
                    "uncommented_changes_reviewed: true",
                    "draft_scientific_paper_skill_used: true",
            ],
            "tone_and_concision": [
                "tone_reviewed: true",
                "redundancy_checked: true",
                "comment_scope_preserved: true",
                "edit_scientific_prose_skill_used: true",
            ],
        }[pass_name]
        report.write_text("\n".join(report_lines), encoding="utf-8")
        return type("Result", (), {"stdout": ""})()

    def mock_run(command, check=True, cwd=None):
        return fake_run(command)

    monkeypatch.setattr("asta_revision_workflow.pandoc_revision_agent_runner.run", mock_run)

    assert run_all(manifest, run_dir, "agent {run_dir} --pass-name {pass_name} --report {report}") == 0
    assert len(calls) == 4
    assert "Implemented." in revised_markdown.read_text(encoding="utf-8")
    audit = json.loads((run_dir / "agent_workflow" / "agent_workflow_audit.json").read_text(encoding="utf-8"))
    assert audit["overall"]["ready_for_finalize"]
    assert all(item["status"] == "completed" for item in audit["passes"])


def test_runner_stops_on_step2_pending_asta_without_resolver(tmp_path, monkeypatch):
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
                            "required_checks": [
                                "revisions_applied",
                                "comment_scope_preserved",
                                "source_docx_only",
                                "draft_scientific_paper_skill_used",
                                "edit_scientific_prose_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper", "edit-scientific-prose"],
                        },
                        {
                            "name": "asta_query_and_collation",
                            "report": "agent_workflow/reports/asta_query_and_collation_report.md",
                            "required_checks": [
                                "modified_claims_citation_checked",
                                "unsupported_claims_resolved",
                                "asta_requests_collated",
                                "source_docx_only",
                                "draft_scientific_paper_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper"],
                        },
                        {
                            "name": "rigor_critique",
                            "report": "agent_workflow/reports/rigor_critique_report.md",
                            "required_checks": [
                                "rigor_approved",
                                "new_knowledge_claims_skeptically_reviewed",
                                "uncommented_changes_reviewed",
                                "draft_scientific_paper_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper"],
                        },
                        {
                            "name": "tone_and_concision",
                            "report": "agent_workflow/reports/tone_concision_report.md",
                            "required_checks": [
                                "tone_reviewed",
                                "redundancy_checked",
                                "comment_scope_preserved",
                                "edit_scientific_prose_skill_used",
                            ],
                            "required_skills": ["edit-scientific-prose"],
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
        "asta_query_and_collation",
        "rigor_critique",
        "tone_and_concision",
    ]:
        (run_dir / "agent_workflow" / "tasks" / f"{name}.md").write_text("# task\n", encoding="utf-8")

    (run_dir / "agent_workflow" / "asta_requests.json").write_text(
        json.dumps(
            {
                "version": 1,
                "requests": [
                    {"id": "yeast-h3k27ac", "required": True, "status": "pending", "claim": "Requires query"},
                ],
            }
        ),
        encoding="utf-8",
    )

    calls: list[list[str]] = []

    def fake_run(command):
        calls.append(command)
        for token in command:
            if token.endswith("revision_implementation_report.md"):
                report = Path(token)
                report_lines = [
                    "revisions_applied: true",
                    "comment_scope_preserved: true",
                    "source_docx_only: true",
                    "draft_scientific_paper_skill_used: true",
                    "edit_scientific_prose_skill_used: true",
                    '```json',
                    '{"markdown_replacements":[],"asta_requests":[]}',
                    "```",
                ]
                report.write_text("\n".join(report_lines), encoding="utf-8")
                return type("Result", (), {"stdout": ""})()
            if token.endswith("asta_query_and_collation_report.md"):
                report = Path(token)
                report_lines = [
                    "modified_claims_citation_checked: false",
                    "unsupported_claims_resolved: false",
                    "asta_requests_collated: false",
                    "source_docx_only: false",
                    "draft_scientific_paper_skill_used: false",
                ]
                report.write_text("\n".join(report_lines), encoding="utf-8")
                return type("Result", (), {"stdout": ""})()
        raise AssertionError("missing report path in fake command")

    def mock_run(command, check=True, cwd=None):
        return fake_run(command)

    monkeypatch.setattr("asta_revision_workflow.pandoc_revision_agent_runner.run", mock_run)
    monkeypatch.delenv("ASTA_REVISION_ASTA_COMMAND", raising=False)

    with pytest.raises(SystemExit, match="no Asta resolver is configured"):
        run_all(manifest, run_dir, "agent {run_dir} --pass-name {pass_name} --report {report}")

    assert len(calls) == 2


def test_agent_runner_stops_step1_if_knowledge_checks_are_declared_without_asta_requests(tmp_path, monkeypatch):
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
                            "required_checks": [
                                "revisions_applied",
                                "comment_scope_preserved",
                                "source_docx_only",
                                "draft_scientific_paper_skill_used",
                                "edit_scientific_prose_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper", "edit-scientific-prose"],
                        },
                        {
                            "name": "asta_query_and_collation",
                            "report": "agent_workflow/reports/asta_query_and_collation_report.md",
                            "required_checks": [
                                "modified_claims_citation_checked",
                                "unsupported_claims_resolved",
                                "asta_requests_collated",
                                "source_docx_only",
                                "draft_scientific_paper_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper"],
                        },
                        {
                            "name": "rigor_critique",
                            "report": "agent_workflow/reports/rigor_critique_report.md",
                            "required_checks": [
                                "rigor_approved",
                                "new_knowledge_claims_skeptically_reviewed",
                                "uncommented_changes_reviewed",
                                "draft_scientific_paper_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper"],
                        },
                        {
                            "name": "tone_and_concision",
                            "report": "agent_workflow/reports/tone_concision_report.md",
                            "required_checks": [
                                "tone_reviewed",
                                "redundancy_checked",
                                "comment_scope_preserved",
                                "edit_scientific_prose_skill_used",
                            ],
                            "required_skills": ["edit-scientific-prose"],
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
        "asta_query_and_collation",
        "rigor_critique",
        "tone_and_concision",
    ]:
        (run_dir / "agent_workflow" / "tasks" / f"{name}.md").write_text("# task\n", encoding="utf-8")

    calls: list[list[str]] = []

    def fake_run(command):
        calls.append(command)
        for token in command:
            if token.endswith("revision_implementation_report.md"):
                report = Path(token)
                report_lines = [
                    "revisions_applied: true",
                    "comment_scope_preserved: true",
                    "source_docx_only: true",
                    "draft_scientific_paper_skill_used: true",
                    "edit_scientific_prose_skill_used: true",
                    '```json',
                    (
                        '{"markdown_replacements":[],'
                        '"required_knowledge_checks":["This sentence lacks citation support"],'
                        '"asta_requests":[]}'
                    ),
                    "```",
                ]
                report.write_text("\n".join(report_lines), encoding="utf-8")
                return type("Result", (), {"stdout": ""})()
        raise AssertionError("missing report path in fake command")

    def mock_run(command, check=True, cwd=None):
        return fake_run(command)

    monkeypatch.setattr("asta_revision_workflow.pandoc_revision_agent_runner.run", mock_run)
    monkeypatch.delenv("ASTA_REVISION_ASTA_COMMAND", raising=False)

    with pytest.raises(SystemExit, match="no required `asta_requests` payload entries"):
        run_all(manifest, run_dir, "agent {run_dir} --pass-name {pass_name} --report {report}")

    assert len(calls) == 1


def test_agent_runner_accepts_four_pass_names_and_computes_overall_readiness(tmp_path, monkeypatch):
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
                            "required_checks": [
                                "revisions_applied",
                                "comment_scope_preserved",
                                "source_docx_only",
                                "draft_scientific_paper_skill_used",
                                "edit_scientific_prose_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper", "edit-scientific-prose"],
                        },
                        {
                            "name": "asta_query_and_collation",
                            "report": "agent_workflow/reports/asta_query_and_collation_report.md",
                            "required_checks": [
                                "modified_claims_citation_checked",
                                "unsupported_claims_resolved",
                                "asta_requests_collated",
                                "source_docx_only",
                                "draft_scientific_paper_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper"],
                        },
                        {
                            "name": "rigor_critique",
                            "report": "agent_workflow/reports/rigor_critique_report.md",
                            "required_checks": [
                                "rigor_approved",
                                "new_knowledge_claims_skeptically_reviewed",
                                "uncommented_changes_reviewed",
                                "draft_scientific_paper_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper"],
                        },
                        {
                            "name": "tone_and_concision",
                            "report": "agent_workflow/reports/tone_concision_report.md",
                            "required_checks": [
                                "tone_reviewed",
                                "redundancy_checked",
                                "comment_scope_preserved",
                                "edit_scientific_prose_skill_used",
                            ],
                            "required_skills": ["edit-scientific-prose"],
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
        "asta_query_and_collation",
        "rigor_critique",
        "tone_and_concision",
    ]:
        (run_dir / "agent_workflow" / "tasks" / f"{name}.md").write_text("# task\n", encoding="utf-8")
    (run_dir / "agent_workflow" / "asta_requests.json").write_text('{"version":1,"requests":[]}', encoding="utf-8")

    def fake_run(command):
        for token in command:
            if token.endswith("revision_implementation_report.md"):
                Path(token).write_text(
                    "revisions_applied: true\ncomment_scope_preserved: true\nsource_docx_only: true\n"
                    "draft_scientific_paper_skill_used: true\nedit_scientific_prose_skill_used: true\n"
                    "```json\n{\"markdown_replacements\":[],\"asta_requests\":[]}\n```",
                    encoding="utf-8",
                )
                return type("Result", (), {"stdout": ""})()
            if token.endswith("asta_query_and_collation_report.md"):
                Path(token).write_text(
                    "modified_claims_citation_checked: true\nunsupported_claims_resolved: true\n"
                    "asta_requests_collated: true\nsource_docx_only: true\ndraft_scientific_paper_skill_used: true",
                    encoding="utf-8",
                )
                return type("Result", (), {"stdout": ""})()
            if token.endswith("rigor_critique_report.md"):
                Path(token).write_text(
                    "rigor_approved: true\nnew_knowledge_claims_skeptically_reviewed: true\n"
                    "uncommented_changes_reviewed: true\ndraft_scientific_paper_skill_used: true",
                    encoding="utf-8",
                )
                return type("Result", (), {"stdout": ""})()
            if token.endswith("tone_concision_report.md"):
                Path(token).write_text(
                    "tone_reviewed: true\nredundancy_checked: true\n"
                    "comment_scope_preserved: true\nedit_scientific_prose_skill_used: true",
                    encoding="utf-8",
                )
                return type("Result", (), {"stdout": ""})()
        raise AssertionError("missing report path in fake command")

    def mock_run(command, check=True, cwd=None):
        return fake_run(command)

    monkeypatch.setattr("asta_revision_workflow.pandoc_revision_agent_runner.run", mock_run)
    monkeypatch.delenv("ASTA_REVISION_ASTA_COMMAND", raising=False)

    assert run_all(manifest, run_dir, "agent {run_dir} --pass-name {pass_name} --report {report}") == 0
    audit = json.loads((run_dir / "agent_workflow" / "agent_workflow_audit.json").read_text(encoding="utf-8"))
    assert len(audit["passes"]) == 4
    assert audit["overall"]["ready_for_finalize"]
    assert audit["overall"]["all_comments_addressed"]


def test_tone_and_concision_pass_applies_json_replacements(tmp_path, monkeypatch):
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
    revised_markdown.write_text("The manuscript needs edits.\n", encoding="utf-8")
    comments_md.write_text("# comments\n", encoding="utf-8")
    comments_json.write_text("[]", encoding="utf-8")
    (run_dir / "comment_scoped_revised.md").write_text("The manuscript needs edits.\n", encoding="utf-8")

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
                "agent_inputs": {"comment_scoped_revised_markdown": "comment_scoped_revised.md"},
                "agent_workflow": {
                    "audit_file": "agent_workflow/agent_workflow_audit.json",
                    "asta_requests": "agent_workflow/asta_requests.json",
                    "required_passes": [
                        {
                            "name": "revision_implementation",
                            "report": "agent_workflow/reports/revision_implementation_report.md",
                            "required_checks": [
                                "revisions_applied",
                                "comment_scope_preserved",
                                "source_docx_only",
                                "draft_scientific_paper_skill_used",
                                "edit_scientific_prose_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper", "edit-scientific-prose"],
                        },
                        {
                            "name": "asta_query_and_collation",
                            "report": "agent_workflow/reports/asta_query_and_collation_report.md",
                            "required_checks": [
                                "modified_claims_citation_checked",
                                "unsupported_claims_resolved",
                                "asta_requests_collated",
                                "source_docx_only",
                                "draft_scientific_paper_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper"],
                        },
                        {
                            "name": "rigor_critique",
                            "report": "agent_workflow/reports/rigor_critique_report.md",
                            "required_checks": [
                                "rigor_approved",
                                "new_knowledge_claims_skeptically_reviewed",
                                "uncommented_changes_reviewed",
                                "draft_scientific_paper_skill_used",
                            ],
                            "required_skills": ["draft-scientific-paper"],
                        },
                        {
                            "name": "tone_and_concision",
                            "report": "agent_workflow/reports/tone_concision_report.md",
                            "required_checks": [
                                "tone_reviewed",
                                "redundancy_checked",
                                "comment_scope_preserved",
                                "edit_scientific_prose_skill_used",
                            ],
                            "required_skills": ["edit-scientific-prose"],
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
        "asta_query_and_collation",
        "rigor_critique",
        "tone_and_concision",
    ]:
        (run_dir / "agent_workflow" / "tasks" / f"{name}.md").write_text("# task\n", encoding="utf-8")
    (run_dir / "agent_workflow" / "asta_requests.json").write_text('{"version":1,"requests":[]}', encoding="utf-8")

    def fake_run(command):
        for token in command:
            if token.endswith("revision_implementation_report.md"):
                Path(token).write_text(
                    "revisions_applied: true\ncomment_scope_preserved: true\nsource_docx_only: true\n"
                    "draft_scientific_paper_skill_used: true\nedit_scientific_prose_skill_used: true\n"
                    "```json\n{\"markdown_replacements\":[],\"asta_requests\":[]}\n```",
                    encoding="utf-8",
                )
                return type("Result", (), {"stdout": ""})()
            if token.endswith("asta_query_and_collation_report.md"):
                Path(token).write_text(
                    "modified_claims_citation_checked: true\nunsupported_claims_resolved: true\n"
                    "asta_requests_collated: true\nsource_docx_only: true\ndraft_scientific_paper_skill_used: true",
                    encoding="utf-8",
                )
                return type("Result", (), {"stdout": ""})()
            if token.endswith("rigor_critique_report.md"):
                Path(token).write_text(
                    "rigor_approved: true\nnew_knowledge_claims_skeptically_reviewed: true\n"
                    "uncommented_changes_reviewed: true\ndraft_scientific_paper_skill_used: true",
                    encoding="utf-8",
                )
                return type("Result", (), {"stdout": ""})()
            if token.endswith("tone_concision_report.md"):
                Path(token).write_text(
                    "\n".join(
                        [
                            "tone_reviewed: true",
                            "redundancy_checked: true",
                            "comment_scope_preserved: true",
                            "edit_scientific_prose_skill_used: true",
                            '```json\n{"markdown_replacements":[{"old":"The manuscript needs edits.","new":"The manuscript needs concise edits."}],"asta_requests":[]}\n```',
                        ]
                    ),
                    encoding="utf-8",
                )
                return type("Result", (), {"stdout": ""})()
        raise AssertionError("missing report path in fake command")

    def mock_run(command, check=True, cwd=None):
        return fake_run(command)

    monkeypatch.setattr("asta_revision_workflow.pandoc_revision_agent_runner.run", mock_run)

    assert run_all(manifest, run_dir, "agent {run_dir} --pass-name {pass_name} --report {report}") == 0
    assert "concise edits" in revised_markdown.read_text(encoding="utf-8")
    assert "concise edits" in (run_dir / "comment_scoped_revised.md").read_text(encoding="utf-8")
