"""Tests for the asta-revision-workflow refactor additions:

- per-pass model sizing in the scaffold
- the claim-level redundancy artifact generator
- bipartite resolver helpers (parse/normalize/dedup building blocks)
- claude-style stdout->report capture in run_single_pass
"""

from pathlib import Path
import json

from asta_revision_workflow.pandoc_revision_launcher import AGENT_WORKFLOW_PASSES, write_agent_workflow_tasks
from asta_revision_workflow.pandoc_revision_agent_runner import (
    run_single_pass,
    write_cite_backed_statements,
)
from asta_revision_workflow.asta_evidence_resolver import (
    normalize_doi,
    paper_doi,
    parse_bip_papers,
    select_papers,
)


# --- per-pass model sizing -------------------------------------------------

def test_passes_declare_per_pass_models():
    models = {p["name"]: p["model"] for p in AGENT_WORKFLOW_PASSES}
    assert models["revision_implementation"] == "claude-opus-4-8"
    assert models["asta_query_and_collation"] == "claude-opus-4-8"
    assert models["rigor_critique"] == "claude-sonnet-4-6"
    assert models["tone_and_concision"] == "claude-sonnet-4-6"


def test_scaffold_threads_model_into_required_passes(tmp_path):
    manifest = {
        "source_sha256": "abc",
        "source_docx": "source.docx",
        "comments": {"markdown": "c.md", "json": "c.json"},
        "citation_policy": {"metadata_overlay_ris": "citation_metadata.ris"},
        "generated_artifacts": {"source_markdown": "s.md", "revised_markdown": "r.md", "media_dir": "media"},
    }
    workflow = write_agent_workflow_tasks(tmp_path, manifest)
    by_name = {p["name"]: p for p in workflow["required_passes"]}
    assert by_name["revision_implementation"]["model"] == "claude-opus-4-8"
    assert by_name["rigor_critique"]["model"] == "claude-sonnet-4-6"
    # claim-level guard check is required of the asta pass
    assert "claim_redundancy_checked" in by_name["asta_query_and_collation"]["required_checks"]


# --- claim-level redundancy artifact ---------------------------------------

def test_cite_backed_statements_lists_only_cited_sentences(tmp_path):
    revised = tmp_path / "r.md"
    revised.write_text(
        "H3K27ac marks active enhancers in yeast ^12^. This sentence has none. "
        "Histone turnover is rapid {Smith, 2020 #44}.\n\nA trailing claim without support.\n",
        encoding="utf-8",
    )
    manifest = {"generated_artifacts": {"revised_markdown": "r.md"}}
    out = write_cite_backed_statements(tmp_path, manifest)
    text = out.read_text(encoding="utf-8")
    assert "H3K27ac marks active enhancers in yeast ^12^." in text
    assert "Histone turnover is rapid {Smith, 2020 #44}." in text
    assert "This sentence has none." not in text
    assert "trailing claim without support" not in text


def test_cite_backed_statements_handles_no_citations(tmp_path):
    revised = tmp_path / "r.md"
    revised.write_text("A claim. Another claim.\n", encoding="utf-8")
    out = write_cite_backed_statements(tmp_path, {"generated_artifacts": {"revised_markdown": "r.md"}})
    assert "No cite-backed statements detected" in out.read_text(encoding="utf-8")


# --- bipartite resolver helpers --------------------------------------------

def test_normalize_doi_strips_prefixes_and_lowercases():
    assert normalize_doi("https://doi.org/10.1/AbC") == "10.1/abc"
    assert normalize_doi("doi:10.2/XY") == "10.2/xy"
    assert normalize_doi("10.3/Zz") == "10.3/zz"


def test_paper_doi_reads_doi_and_external_ids():
    assert paper_doi({"doi": "10.1/a"}) == "10.1/a"
    assert paper_doi({"externalIds": {"DOI": "10.2/B"}}) == "10.2/b"
    assert paper_doi({"externalIds": None}) == ""


def test_parse_bip_papers_accepts_list_object_and_jsonl():
    one = '[{"title":"A","year":2020,"authors":["Doe, J"],"doi":"10.1/a"}]'
    two = '{"results":[{"title":"B","year":2021,"authors":["Roe, K"],"doi":"10.2/b"}]}'
    three = '{"title":"C","year":2022,"authors":["Lee, M"]}\n{"title":"D","year":2023,"authors":["Ng, P"]}'
    assert len(parse_bip_papers(one)) == 1
    assert len(parse_bip_papers(two)) == 1
    assert len(parse_bip_papers(three)) == 2
    assert parse_bip_papers("") == []


def test_reference_level_dedup_filters_nexus_dois():
    candidates = parse_bip_papers(
        '[{"title":"A","year":2020,"authors":["Doe, J"],"doi":"10.1/a"},'
        '{"title":"B","year":2021,"authors":["Roe, K"],"doi":"10.2/B"}]'
    )
    nexus = {"10.1/a"}  # paper A already in the bibliography
    survivors = [p for p in candidates if paper_doi(p) not in nexus or not paper_doi(p)]
    assert len(survivors) == 1
    assert paper_doi(survivors[0]) == "10.2/b"


def test_select_papers_rejects_incomplete_records():
    papers = parse_bip_papers(
        '[{"title":"Complete","year":2020,"authors":["Doe, J"]},'
        '{"title":"NoYear","authors":["Roe, K"]},'
        '{"title":"NoAuthors","year":2021,"authors":[]}]'
    )
    selected = select_papers(papers, 5)
    assert [p["title"] for p in selected] == ["Complete"]


# --- claude-style stdout -> report capture ---------------------------------

def _minimal_manifest(tmp_path):
    (tmp_path / "source.docx").write_text("", encoding="utf-8")
    (tmp_path / "s.md").write_text("# source\n", encoding="utf-8")
    (tmp_path / "r.md").write_text("# revised\n", encoding="utf-8")
    (tmp_path / "c.md").write_text("# comments\n", encoding="utf-8")
    (tmp_path / "c.json").write_text("[]", encoding="utf-8")
    (tmp_path / "agent_workflow" / "tasks").mkdir(parents=True)
    (tmp_path / "agent_workflow" / "tasks" / "tone_and_concision.md").write_text("# task\n", encoding="utf-8")
    (tmp_path / "agent_workflow" / "asta_requests.json").write_text('{"version":1,"requests":[]}', encoding="utf-8")
    return {
        "source_docx": "source.docx",
        "source_sha256": "abc",
        "generated_artifacts": {"source_markdown": "s.md", "revised_markdown": "r.md", "media_dir": "media"},
        "comments": {"markdown": "c.md", "json": "c.json"},
        "agent_workflow": {
            "audit_file": "agent_workflow/agent_workflow_audit.json",
            "asta_requests": "agent_workflow/asta_requests.json",
        },
    }


def test_run_single_pass_captures_stdout_when_report_untouched(tmp_path, monkeypatch):
    manifest = _minimal_manifest(tmp_path)
    manifest_path = tmp_path / "manifest.json"
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")

    stdout_report = "\n".join(
        [
            "# tone_and_concision",
            "tone_reviewed: true",
            "redundancy_checked: true",
            "comment_scope_preserved: true",
            "edit_scientific_prose_skill_used: true",
        ]
    )

    # Simulate `claude -p`: prints the report to stdout, never writes {report}.
    def fake_run(command, check=True, cwd=None):
        assert cwd == tmp_path  # runner sets cwd=run_dir for skill discovery
        return type("Result", (), {"stdout": stdout_report})()

    monkeypatch.setattr("asta_revision_workflow.pandoc_revision_agent_runner.run", fake_run)

    pass_definition = {
        "name": "tone_and_concision",
        "report": "agent_workflow/reports/tone_concision_report.md",
        "model": "claude-sonnet-4-6",
        "required_checks": [
            "tone_reviewed",
            "redundancy_checked",
            "comment_scope_preserved",
            "edit_scientific_prose_skill_used",
        ],
    }
    result = run_single_pass(
        tmp_path, manifest_path, manifest, pass_definition, "claude -p --model {model} --add-dir {run_dir}"
    )
    assert result["status"] == "completed"
    report = (tmp_path / pass_definition["report"]).read_text(encoding="utf-8")
    assert report == stdout_report
