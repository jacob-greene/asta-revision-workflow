# Agentic Revision Workflow

This workflow revises a commented Word draft while preserving content provenance and citation reproducibility. It is designed for scientific manuscripts, thesis chapters, grants, and reports.

The current commented `.docx` is the source of truth for each revision round. Do not restore text from archived drafts, Markdown exports, or TeX sources unless those files were generated from the same commented `.docx` in the current pass.

Each pass should produce a paired `.docx` and `.ris` with the same stem:

```text
manuscript_v4.docx
manuscript_v4.ris
```

## Inputs

- Exactly one content input: the current commented Word draft.
- Comment extraction generated from that same Word draft by `word-doc-only-revision start`.
- Do not use archived DOCX files, Markdown exports, TeX sources, response files, cached Asta evidence, external RIS files, or BibTeX files as content or citation sources for the revision pass.
- The paired RIS must be generated only from the reference list in the same revised Word document produced during the pass.
- The scientific-writing skills in `codex-skills/`, especially `draft-scientific-paper` and `edit-scientific-prose`.

## Comment Extraction

Launch the pass from the current Word document before planning edits:

```bash
word-doc-only-revision start commented-draft.docx --output-stem manuscript_v4
```

This command creates a run directory, copies the source Word document into it, records the source hash, extracts Word comments, and writes a manifest naming the only permitted generated artifacts for that pass.

Do not use Pandoc, Markdown export, plain text extraction, archive RIS files, prior response files, or cached evidence as the comment or content source. Those routes can omit Word comments, make an actively commented draft appear un-commented, or restore stale prose.

## Agent Roles

Every revision pass must run all four roles below in order. Do not skip directly from comment extraction to implementation, even when the requested change looks mechanical or citation-only. If a pass starts from a draft with no active Word comments, define the revision scope from the user's request, the previous comment-response summary, and the paragraphs changed in the prior pass; then run the same four roles on that explicit scope.

1. Comment interpretation and revision planning agent
   - Read the `docx-extract-comments` output for every Word comment and its surrounding paragraph.
   - Produce a concrete plan keyed to comment IDs.
   - Produce a current outline based only on the commented draft text.
   - List the exact paragraphs that may be revised.

2. Evidence and specificity agent
   - Scan the full draft for vague, unsupported, or weakly justified claims, focusing on commented regions.
   - For every modified claim, first check whether the same sentence or an adjacent sentence has a citation that plausibly supports it.
   - If nearby existing citations do not support the modified claim, soften or remove the claim within the current Word-doc evidence base.
   - Do not read cached Asta output, prior evidence JSON, prior response files, archive RIS files, or older drafts during the revision pass.
   - Query literature tools or Asta only if the user explicitly authorizes a separate evidence-gathering pass; any resulting reference must be inserted into the current Word document before it can be used by the Word-doc-only revision pass.
   - Recommend citations, caveats, or softer wording.

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

- Start every Word-based revision with `word-doc-only-revision start SOURCE.docx`. This is the only supported launcher for commented Word drafts.
- Do not start from a hand-run rebuild script for a prior version. Version-specific rebuild scripts may be used only after they have been regenerated from the launcher manifest and the current Word document, and only for the paragraph scope in that manifest.
- Patch only commented sections and required adjacent text.
- Do not revise un-commented paragraphs by default.
- The only exception is adjacent text explicitly required by a section-level comment; the plan must identify that exception before implementation.
- Within a commented paragraph, preserve source sentences that are not implicated by the comment. Full-paragraph rewrites are allowed only when the comment asks for paragraph-level restructuring or the revision plan explicitly justifies the rewrite.
- Prefer precise caveats over unsupported citations.
- Do not introduce broad synthesis claims during implementation unless the rigor critique step explicitly approves the claim and identifies the supporting citations.
- Do not rebuild prose from older TeX or Markdown. If EndNote temporary citations are required, first synchronize the exact revised `.docx` content into a fresh citation source.
- Every pass must emit a revised `.docx` and a matching `.ris` generated from the full numbered reference list in the current source Word document for that pass. The RIS must include every numbered reference in that Word bibliography, not only cited records, and the exporter must fail rather than silently skip malformed entries.
- Do not merge, backfill, or repair the pass RIS from an archive RIS, another pass, BibTeX, EndNote library, or web/Asta lookup. If the current Word reference list lacks required metadata, fix the reference entry in the Word document itself before launching the workflow, then regenerate from that Word document.
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

When a revision pass starts from a Word draft with numeric superscript citations and a numbered reference list, use the launcher. It accepts only the source Word document as the external content input:

```bash
word-doc-only-revision start commented-draft.docx --output-stem manuscript_v4
```

After the revision agent creates the manifest-named raw revised DOCX inside the run directory, finalize the pass:

```bash
word-doc-only-revision finalize manuscript_v4_word_doc_only_run/manifest.json
```

The finalizer runs:

```text
docx-reference-list-to-ris
docx-modified-citation-support
docx-plain-numeric-citation-check
docx-numeric-to-endnote-temp
docx-word-sanity
docx-endnote-ris-sync
```

`docx-reference-list-to-ris` uses the current source Word reference list as the only citation authority. The launcher validates that the paired RIS is exactly the canonical RIS derived from that source Word document, so archive-filled metadata and older records fail provenance checks.

`docx-numeric-to-endnote-temp` converts numeric superscript citations into EndNote temporary citations. Pass `--ris` so citation metadata comes from the complete paired RIS rather than a lossy Word reference parse. When two distinct references share the same first author and year, the temporary citation includes the title so EndNote matching is unique. Duplicate entries for the same paper remain concise.

By default, the static reference list is removed from the generated `.docx`; EndNote should regenerate the bibliography after the user imports the paired RIS and runs `Update Citations and Bibliography`.

## Citation Pipeline for TeX Drafts

When a revision pass starts from TeX/BibTeX:

```bash
CITEPROC_ENDNOTE_BIBLIOGRAPHY=references.bib \
  pandoc -f latex -t docx input.tex -o manuscript_v4.docx -F citeproc-endnote

bib-to-ris references.bib manuscript_v4.ris
docx-word-sanity manuscript_v4.docx
```

Use this route only when the TeX source is the active source of truth for that revision pass.

## Finish in EndNote

1. Import the same-stem `.ris` into EndNote.
2. Open the generated `.docx` in Word.
3. Run EndNote's `Update Citations and Bibliography`.

The generated `.docx` contains EndNote temporary citations, not final EndNote fields. EndNote performs the field conversion inside Word.
