# Agentic Revision Workflow

This workflow revises a commented Word draft while preserving content provenance and citation reproducibility. It is designed for scientific manuscripts, thesis chapters, grants, and reports.

The current commented `.docx` is the source of truth for each revision round. Do not restore text from archived drafts, Markdown exports, or TeX sources unless those files were generated from the same commented `.docx` in the current pass.

Each pass should produce a paired `.docx` and `.ris` with the same stem:

```text
manuscript_v4.docx
manuscript_v4.ris
```

## Inputs

- Commented Word draft.
- Comment extraction generated from the same Word draft using `docx-extract-comments`.
- Any project-specific evidence files or literature search outputs needed to answer comments.
- Optional BibTeX source when drafting from TeX.
- The scientific-writing skills in `codex-skills/`, especially `draft-scientific-paper` and `edit-scientific-prose`.

## Comment Extraction

Extract Word comments from the DOCX package before planning edits:

```bash
docx-extract-comments commented-draft.docx -o comments.md --format markdown
```

This command reads `word/comments.xml` and the comment anchors in `word/document.xml`. It records the comment ID, comment text, anchored text, and full paragraph text.

Do not use Pandoc, Markdown export, or plain text extraction as the comment source. Those routes can omit Word comments and make an actively commented draft appear un-commented.

## Agent Roles

1. Comment interpretation and revision planning agent
   - Read the `docx-extract-comments` output for every Word comment and its surrounding paragraph.
   - Produce a concrete plan keyed to comment IDs.
   - Produce a current outline based only on the commented draft text.
   - List the exact paragraphs that may be revised.

2. Evidence and specificity agent
   - Scan the full draft for vague, unsupported, or weakly justified claims, focusing on commented regions.
   - Query literature tools only for targeted evidence gaps.
   - Recommend citations, caveats, or softer wording.

3. Rigor critique agent
   - Critique the plan for overclaiming, missing caveats, and logical inconsistency.
   - Flag places where restructuring could accidentally alter uncommented material.
   - Reject edits to un-commented paragraphs unless the plan names the comment that requires the adjacent change.

4. Tone and concision agent
   - Enforce direct topic sentences, short paragraphs, and sentence-to-sentence flow.
   - Preserve scientific nuance while removing clutter.

## Implementation Rules

- Patch only commented sections and required adjacent text.
- Do not revise un-commented paragraphs by default.
- The only exception is adjacent text explicitly required by a section-level comment; the plan must identify that exception before implementation.
- Prefer precise caveats over unsupported citations.
- Keep paragraph length appropriate for the document; for dense scientific prose, 4-5 sentences is a useful default ceiling.
- Before delivery, state which commented sections were revised and whether any un-commented adjacent paragraphs changed under a section-level comment.

## Citation Pipeline for Word Drafts

When a revision pass starts from a Word draft with numeric superscript citations and a numbered reference list, generate the RIS from the revised Word reference list:

```bash
docx-extract-comments commented-draft.docx -o comments.md --format markdown
docx-reference-list-to-ris manuscript_v4_with_refs.docx manuscript_v4.ris
docx-numeric-to-endnote-temp manuscript_v4_with_refs.docx manuscript_v4.docx
docx-word-sanity manuscript_v4.docx
```

`docx-reference-list-to-ris` uses the Word reference list as the citation authority. This ensures references manually added by a collaborator in Word propagate into the RIS.

`docx-numeric-to-endnote-temp` converts numeric superscript citations into EndNote temporary citations. When two distinct references share the same first author and year, the temporary citation includes the title so EndNote matching is unique. Duplicate entries for the same paper remain concise.

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
