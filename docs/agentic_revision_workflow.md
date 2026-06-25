# Agentic Revision Workflow

This workflow revises a commented Word draft while preserving content provenance and citation reproducibility. It is designed for scientific manuscripts, thesis chapters, grants, and reports.

The current commented `.docx` is the source of truth for each revision round. The revision surface is the Pandoc markdown generated from that same `.docx` inside the current run directory. Do not restore text from archived drafts, older Markdown exports, or TeX sources unless those files were generated from the same commented `.docx` in the current pass.

Each pass should produce a paired `.docx` and `.ris` with the same stem:

```text
manuscript_v4.docx
manuscript_v4.ris
```

## Inputs

- Exactly one content input: the current commented Word draft.
- Text/style extraction generated from that Word draft by `asta-revision start`.
- Comment extraction generated from that same Word draft by `asta-revision start`.
- Complete citation metadata is extracted from embedded EndNote fields in the current source DOCX into `citation_metadata.ris`. A `--metadata-ris` file is only a fallback for documents with no embedded EndNote records.
- Do not use archived DOCX files, older Markdown exports, TeX sources, response files, cached Asta evidence, or unpinned external citation files as content or citation sources for the revision pass.
- The paired RIS must be generated from the reference list in the same revised Word document produced during the pass, optionally enriched by the run-local pinned metadata RIS and explicitly recorded Asta additions.
- The scientific-writing skills in `skills/`, especially `draft-scientific-paper` and `edit-scientific-prose`.

## Comment Extraction

Launch the pass from the current Word document before planning edits:

```bash
asta-revision start commented-draft.docx \
  --output-stem manuscript_v4
```

This command creates a run directory, copies the source Word document into it, records the source hash, extracts Word comments, exports the document text/styles to Pandoc markdown, saves `style-reference.docx`, extracts embedded EndNote metadata into `citation_metadata.ris`, creates `agent_inputs/` scoped context files, creates `agent_workflow/` task files plus an audit template, writes `launcher_profile.json`, and writes a manifest naming the permitted generated artifacts for that pass.

Do not use hand-created Markdown exports, archive RIS files, prior response files, or cached evidence as the comment or content source. The only supported Markdown source is the run-local markdown produced by `asta-revision start` from the current `.docx`.

## Agent Roles

Every revision pass must run all four roles below in order. Do not skip directly from comment extraction to implementation, even when the requested change looks mechanical or citation-only. If a pass starts from a draft with no active Word comments, define the revision scope from the user's request, the previous comment-response summary, and the paragraphs changed in the prior pass; then run the same four roles on that explicit scope.

The Pandoc launcher enforces this. Before finalization, `agent_workflow/agent_workflow_audit.json` must exist, must hash the exact `*.revised.md` being finalized, must mark all four passes as completed, and must point to non-empty report files in `agent_workflow/reports/`.

For efficiency, each task file names recommended minimal inputs from `agent_inputs/agent_input_manifest.json`. Revision implementation should start from comment-scoped files, and rigorous/concision passes should avoid loading the full `citation_metadata.ris` unless explicitly needed. The Asta query-and-collation pass may load full citation context because evidence collation is its role.

1. Revision implementation agent
   - Read the `*.comments.md` and `*.comments.json` outputs for every Word comment and its surrounding paragraph.
   - Read the run-local `*.source.md` before editing `*.revised.md`.
   - Implement scoped edits in `*.revised.md` from the current Word-derived source.
   - Preserve scope; do not revise un-commented text except explicit adjacent exceptions.
   - Convert any retained unsupported claims into pending Asta requests when citation support is missing.

2. Asta query-and-collation agent
   - Read `agent_workflow/asta_requests.json` and resolve required requests that the earlier pass emitted.
   - Verify Asta outputs are produced and collated in run-local artifacts.
   - Update the claim-level evidence status: supported claims, softened claims, or unresolved requests.
   - Confirm unresolved claims are not left in `*.revised.md` without a documented justification.

3. Rigor critique agent
   - Critique the plan for overclaiming, missing caveats, and logical inconsistency.
   - Treat every new knowledge claim as provisional until tied to specific evidence. Be highly skeptical of broad, causal, conserved, universal, or mechanistic claims introduced during revision.
   - Prefer narrower wording over adding a citation to an overbroad statement. If the cited evidence is indirect, model-based, organism-specific, or context-specific, the sentence must say so.
   - Flag places where restructuring could accidentally alter uncommented material.
   - Reject edits to un-commented paragraphs unless the plan names the comment that requires the adjacent change.

4. Tone and concision agent
   - Enforce direct topic sentences, short paragraphs, and sentence-to-sentence flow.
   - Preserve scientific nuance while removing clutter.

## Implementation Rules

- Start every Word-based revision with `asta-revision start SOURCE.docx`. This is the supported launcher for commented Word drafts.
- Do not start from a hand-run rebuild script for a prior version. Version-specific rebuild scripts may be used only after they have been regenerated from the launcher manifest and the current Word document, and only for the paragraph scope in that manifest.
- Patch only commented sections and required adjacent text.
- Do not revise un-commented paragraphs by default.
- The only exception is adjacent text explicitly required by a section-level comment; the plan must identify that exception before implementation.
- Within a commented paragraph, preserve source sentences that are not implicated by the comment. Full-paragraph rewrites are allowed only when the comment asks for paragraph-level restructuring or the revision plan explicitly justifies the rewrite.
- Prefer precise caveats over unsupported citations.
- Do not introduce broad synthesis claims during implementation unless the rigor critique step explicitly approves the claim and identifies the supporting citations.
- Do not rebuild prose from older TeX or Markdown. The active prose source after launch is the run-local revised markdown generated from the current Word document.
- Every pass must emit a revised `.docx` and a matching `.ris` generated from the full numbered reference list in the current recompiled Word document for that pass. The RIS must include every numbered reference in that Word bibliography, not only cited records, and the exporter must fail rather than silently skip malformed entries.
- Do not merge, backfill, or repair the pass RIS from an archive RIS, another pass, BibTeX, EndNote library, or web/Asta lookup. Complete metadata should come from embedded EndNote records in the current source DOCX. New Asta citations must be recorded explicitly in `asta_reference_additions.json` and inserted into the current revised markdown/reference list.
- If the user requests a writing-methods or AI-use statement, append it after the numbered bibliography as a non-reference methods note and link to the run manifest or workflow documentation. The RIS exporter must stop at the end of the numbered bibliography so this statement cannot be absorbed into the last reference.
- Keep paragraph length appropriate for the document; for dense scientific prose, 4-5 sentences is a useful default ceiling.
- Before converting numeric citations to EndNote temporary citations, run the modified-statement support check on the revised raw DOCX. Any modified sentence without a same-sentence or adjacent citation must be resolved by checking nearby existing citations, adding a citation, softening/removing the claim, or requerying literature tools for targeted evidence.
- Also before conversion, run the plain numeric citation check on the revised raw DOCX. This catches intake failures where citations next to digit-containing scientific terms, such as `PRC2`, `H3K27me3`, or `H3.3K27M`, were not converted into superscript citation runs.
- After converting to EndNote temporary citations, run the DOCX/RIS sync check. There must be no `REF#` placeholders, missing RIS entries, or ambiguous author-year temporary citations without title disambiguation.
- If Word or EndNote crashes while updating citations, generate an EndNote-safe copy with short title-prefix disambiguation for duplicate author-year groups. Prefer `{Author, Year, First few title words}` over full-title temporary citations; this keeps records unique without long comma-heavy strings that can destabilize EndNote.
- EndNote-ready DOCX files must not mix new temporary citations with old formatted EndNote fields. Before delivery, strip `ADDIN EN.CITE` field wrappers, stale `EN.*` document variables, comment parts, and unused EndNote bibliography styles; then verify those strings are absent from the final DOCX package.
- Before delivery, compare the revised raw DOCX against the current source Word draft. All un-commented paragraphs must match exactly, and comment-scoped paragraphs must be reviewed for unintended reversion to older wording or older citation numbering.
- Before delivery, state which commented sections were revised and whether any un-commented adjacent paragraphs changed under a section-level comment.

## Citation Pipeline for Word Drafts

When a revision pass starts from a Word draft with numeric superscript citations and a numbered reference list, use the Pandoc launcher. It accepts the current source Word document as the content input and may accept a pinned complete RIS metadata overlay:

```bash
asta-revision start commented-draft.docx \
  --output-stem manuscript_v4
```

Revise only the run-local markdown named in the manifest:

```bash
manuscript_v4_asta_revision_run/manuscript_v4.revised.md
```

Then finalize the pass:

```bash
asta-revision finalize manuscript_v4_asta_revision_run/manifest.json
```

Finalization is blocked until the agent workflow audit is complete:

```text
agent_workflow/agent_workflow_audit.json
agent_workflow/reports/asta_query_and_collation_report.md
agent_workflow/reports/rigor_critique_report.md
agent_workflow/reports/tone_concision_report.md
```

Launcher timing and token estimates are written to:

```text
launcher_profile.json
agent_inputs/agent_input_manifest.json
agent_inputs/comment_scoped_source.md
agent_inputs/comment_scoped_revised.md
```

The finalizer runs:

```text
pandoc markdown -> docx using style-reference.docx
docx-reference-list-to-ris
docx-plain-numeric-citation-check
docx-numeric-to-endnote-temp
docx-word-sanity
docx-endnote-ris-sync
deterministic repeated citation conversion check
```

`docx-reference-list-to-ris` uses the current recompiled Word reference list as the citation membership/order authority. If `citation_metadata.ris` is present in the run directory, it supplies complete author, DOI, journal, volume, and page fields for matching titles only; it cannot add references that are absent from the current Word document.

`docx-numeric-to-endnote-temp` converts numeric superscript citations into EndNote temporary citations. Pass `--ris` so citation metadata comes from the complete paired RIS rather than a lossy Word reference parse. When two distinct references share the same first author and year, the temporary citation includes the title so EndNote matching is unique. Duplicate entries for the same paper remain concise.

By default, the static reference list is removed from the generated `.docx`; EndNote should regenerate the bibliography after the user imports the paired RIS and runs `Update Citations and Bibliography`.

## Finish in EndNote

1. Import the same-stem `.ris` into EndNote.
2. Open the generated `.docx` in Word.
3. Run EndNote's `Update Citations and Bibliography`.

The generated `.docx` contains EndNote temporary citations, not final EndNote fields. EndNote performs the field conversion inside Word.
