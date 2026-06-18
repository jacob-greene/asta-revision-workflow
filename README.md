# scientific-writing-endnote-workflow

This repository combines three reusable pieces for scientific writing projects:

- `citeproc-endnote`: a small Pandoc JSON filter that converts LaTeX/BibTeX citations into EndNote temporary citations in editable Word documents.
- Word citation utilities for revision workflows that start from `.docx` drafts with numbered citations and reference lists.
- Codex scientific-writing skills for drafting, editing, reviewing figures, and grant writing.

The tools are project-neutral. A thesis, manuscript, grant, or report can keep its own source files elsewhere and use this repository as the installed workflow layer.

## Typical Word Comment Workflow

The common use case is a Word-first revision round:

1. Write or revise the draft in Word.
2. Add Word comments to the passages that need revision.
3. Submit the commented `.docx` to an agent with the instruction that the commented draft is the source of truth.
4. The agent interprets comments, revises only comment-scoped text, and preserves un-commented sections unless an adjacent change is explicitly justified.
5. The agent returns an updated `.docx` plus a same-stem `.ris`.
6. Import the `.ris` into EndNote, open the `.docx` in Word, and run EndNote's **Update Citations and Bibliography** command.

Simple workflow chart:

```text
Commented Word draft
        |
        v
pandoc-word-revision start
        |
        +--> source markdown
        +--> comments markdown/json
        +--> style-reference.docx
        +--> citation_metadata.ris extracted from embedded EndNote fields
        +--> agent_inputs scoped context and token profile
        +--> agent_workflow tasks and audit template
        |
        v
Agent revision workflow over run-local markdown
        |
        +--> comment plan and outline
        +--> evidence and specificity check
        +--> rigor critique
        +--> tone and concision pass
        |
        v
Pandoc recompiled Word draft with EndNote temporary citations
        |
        +--> paired RIS generated from current references plus embedded metadata
        |
        v
Import RIS into EndNote, then update citations in Word
```

The agent workflow has four explicit passes:

1. **Comment interpretation and revision planning**: read each Word comment with its surrounding text, create a concrete revision plan, and identify exactly which paragraphs may change.
2. **Evidence and specificity**: check the commented regions for vague, unsupported, or overly broad claims; add targeted evidence or soften the wording.
3. **Rigor critique**: review the proposed changes for scientific accuracy, overclaiming, missing caveats, and accidental changes to un-commented sections.
4. **Tone and concision**: make the prose direct, readable, and concise while preserving the scientific meaning.

`pandoc-word-revision finalize` enforces this workflow. It refuses to compile
the final Word/RIS outputs until `agent_workflow/agent_workflow_audit.json`
exists, names the same source DOCX hash, hashes the exact `*.revised.md` being
finalized, marks all four passes complete, points to non-empty pass reports, and
sets the required overall readiness checks to true.

The final implementation should be based only on the provided commented draft and its comments. Do not rebuild from stale Markdown, TeX, or archived Word drafts unless those files were generated from the same current `.docx`.

Launch the Pandoc-centered Word workflow before planning revisions:

```bash
pandoc-word-revision start commented-draft.docx \
  --output-stem manuscript_v4
```

This creates a run directory containing `manuscript_v4.source.md`,
`manuscript_v4.revised.md`, `manuscript_v4.comments.md`,
`manuscript_v4.comments.json`, `style-reference.docx`,
`citation_metadata.ris`, `citation_metadata_audit.json`, `agent_workflow/`, and
a manifest. It also writes `agent_inputs/` and `launcher_profile.json` so each
agent pass can load a smaller, role-appropriate context. Word comments and
complete citation metadata are extracted directly from the DOCX; Pandoc supplies
the editable markdown and the Word style reference. The `agent_workflow/`
directory contains task files, report paths, and an audit template that must be
completed before finalization.

## Install

From this repository:

```bash
uv tool install .
```

From GitHub:

```bash
uv tool install git+https://github.com/jacob-greene/citeproc-endnote-uv
```

For project-local development:

```bash
uv venv
uv pip install -e .
```

## LaTeX to Word with EndNote temporary citations

Pass a BibTeX file to the Pandoc filter explicitly:

```bash
CITEPROC_ENDNOTE_BIBLIOGRAPHY=references.bib \
  pandoc -f latex -t docx input.tex -o output.docx -F citeproc-endnote
```

The generated DOCX contains editable temporary citations such as `{Cao, 2004}`. Open the file in Word, import the matching RIS file into EndNote, then run EndNote's **Update Citations and Bibliography** command.

Convert the same BibTeX file to RIS:

```bash
bib-to-ris references.bib references.ris
```

## Word draft to EndNote-ready Word draft

For workflows that begin from a `.docx` draft with numeric superscript citations and a numbered reference list, use the Pandoc launcher:

```bash
pandoc-word-revision start commented-draft.docx \
  --output-stem manuscript_v4

# edit manuscript_v4_pandoc_revision_run/manuscript_v4.revised.md

pandoc-word-revision finalize manuscript_v4_pandoc_revision_run/manifest.json
```

`pandoc-word-revision start` extracts text and style through Pandoc, extracts
comments directly from Word XML, and extracts complete citation metadata from
embedded EndNote field records in the same source DOCX. `finalize` recompiles
the revised markdown with Pandoc only after validating the required agent
workflow audit. It then generates the paired RIS from the current recompiled
reference list, uses the run-local metadata overlay to restore complete
author/DOI fields, converts numeric citations to EndNote temporary citations,
and runs sanity/sync checks.

`launcher_profile.json` records per-step launcher timings, artifact sizes, and
approximate token counts. `agent_inputs/agent_input_manifest.json` names the
minimal recommended input files for each agent pass. Non-citation passes avoid
loading `citation_metadata.ris` by default, and comment-scoped input files let
planning, rigor, and tone reviewers inspect only the Word-comment anchors unless
they need the full markdown.

Citation handling is deterministic for a fixed run directory. The recompiled
Word document defines which references exist and their numbering/order. The
run-local `citation_metadata.ris` is built from embedded EndNote records in the
current source DOCX, including full author lists even when the visible Word
bibliography says `et al.`. During finalization, abbreviated visible references
must match the run-local metadata; otherwise the workflow fails rather than
emitting truncated authors.

If the evidence pass finds a modified claim that is not supported by adjacent
citations, it must prioritize Asta for claims that should be retained: add a
required item to `agent_workflow/asta_requests.json`. Remove or soften the claim
only when the claim should not remain even with new evidence.
`pandoc-word-revision finalize` resolves pending required requests before
citation generation. Configure the resolver with `--asta-command` or
`PANDOC_REVISION_ASTA_COMMAND`. The command may use placeholders:

```bash
pandoc-word-revision finalize manuscript_run/manifest.json \
  --asta-command 'asta-evidence-resolver --request {request_json} --output {output_json} --ris {output_ris}'
```

If no placeholders are present, the launcher appends `--request`, `--output`,
and `--ris`. The resolver must write complete RIS records with at least type,
title, author, and year. Finalize validates those records, writes
`asta_reference_additions.json`, and combines them with the embedded metadata as
`citation_metadata.with_asta.ris`. If Asta is required but unresolved, finalize
fails and reports the pending request IDs.

`docx-reference-list-to-ris` remains available as a lower-level utility. It
reads the numbered reference list from a Word document and writes a matching RIS
file; pass `--metadata-ris --require-metadata-match` when the Word bibliography
contains `et al.` or otherwise abbreviated metadata.

`docx-numeric-to-endnote-temp` converts numeric superscript citations such as `6,7` into EndNote temporary citations such as `{McCabe, 2012, Mutation of A677...; McCabe, 2012, EZH2 inhibition...}`. When two distinct papers share the same first author and year, the converter includes the title to make EndNote matching unique. Duplicate bibliography entries for the same paper remain concise.

By default, `docx-numeric-to-endnote-temp` removes the static reference list from the output DOCX so EndNote can regenerate the bibliography after updating citations. Use `--keep-references` if you need an inspection copy.

## EndNote import

In EndNote:

1. Open the target library.
2. Choose `File > Import > File...`.
3. Select the generated `.ris` file.
4. Set `Import Option` to `Reference Manager (RIS)` or `RIS`.
5. Import the records.
6. Open the generated DOCX in Word and run EndNote's **Update Citations and Bibliography** command.

These tools emit EndNote temporary citations, not final EndNote fields. EndNote performs the field conversion inside Word.

## Agentic revision workflow

The general workflow is documented in [docs/agentic_revision_workflow.md](docs/agentic_revision_workflow.md). The core rule is that the current commented `.docx` is the source of truth: revisions should be made only in comment-scoped regions unless the review plan explicitly justifies adjacent changes.

## Codex skills

Scientific-writing skills live in [codex-skills/](codex-skills/). Symlink the skills you want into your Codex skills directory, for example:

```bash
ln -s "$PWD/codex-skills/edit-scientific-prose" ~/.codex/skills/edit-scientific-prose
ln -s "$PWD/codex-skills/draft-scientific-paper" ~/.codex/skills/draft-scientific-paper
```

The skills are independent of the citation tools but are bundled here so a project can use one repository for prose guidance, citation conversion, and reproducible revision workflow conventions.
