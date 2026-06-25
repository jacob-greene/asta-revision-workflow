# asta-revision-workflow

A Claude-driven Word revision workflow for scientific manuscripts. It ingests a
commented `.docx`, runs a four-pass agent workflow that addresses the comments
(scientific-writing skills + bipartite-backed literature evidence), and recompiles
a `.docx` whose citations are EndNote temporary citations paired with a same-stem
`.ris` for EndNote's **Update Citations and Bibliography**.

This repository combines:

- Word citation utilities for revision workflows that start from `.docx` drafts
  with numbered citations and reference lists.
- A four-pass agent workflow runner that fans out to headless Claude (`claude -p`).
- A bipartite-backed evidence resolver (search + DOI dedup against a shared
  bibliography) with a claim-level redundancy guard.
- Scientific-writing skills for drafting, editing, evidence collation, figures,
  and grants.

The tools are project-neutral: a thesis, manuscript, grant, or report can keep
its own source files elsewhere and use this repository as the installed workflow
layer.

> The legacy LaTeX/BibTeX → EndNote Pandoc filter (`citeproc-endnote`) and the
> `bib-to-ris` helper have been moved to [`archive/`](archive/); they are no
> longer part of this workflow or installed as commands.

## Typical Word Comment Workflow

1. Write or revise the draft in Word.
2. Add Word comments to the passages that need revision.
3. Run `asta-revision run commented-draft.docx`. The commented draft is the
   source of truth.
4. The workflow interprets comments, revises only comment-scoped text, resolves
   evidence for retained claims, and preserves un-commented sections unless an
   adjacent change is explicitly justified.
5. It returns an updated `.docx` plus a same-stem `.ris`.
6. Import the `.ris` into EndNote, open the `.docx` in Word, and run EndNote's
   **Update Citations and Bibliography** command.

Simple workflow chart:

```text
Commented Word draft
        |
        v
asta-revision run
        |
        +--> source markdown
        +--> comments markdown/json
        +--> style-reference.docx
        +--> citation_metadata.ris extracted from embedded EndNote fields
        +--> agent_inputs scoped context and token profile
        +--> agent_workflow tasks and audit template
        +--> run-local .claude/skills (repo skills, for headless Claude)
        |
        v
Four-pass Claude revision workflow over run-local markdown
        |
        +--> revision implementation
        +--> asta query and collation (bipartite evidence + claim-redundancy guard)
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

1. **Revision implementation** (`claude-opus-4-8`): make scoped paragraph
   replacements from the current Word-derived markdown and record evidence
   requests for unsupported claims that should be retained. Requires
   `draft-scientific-paper` and `edit-scientific-prose`.
2. **Asta query and collation** (`claude-opus-4-8`): apply the claim-level
   redundancy guard, resolve required evidence requests via bipartite, then
   collate returned evidence into run-local artifacts before reviewer passes.
   Requires `draft-scientific-paper` and `asta-query-and-collation`.
3. **Rigor critique** (`claude-sonnet-4-6`): review the proposed changes for
   scientific accuracy, overclaiming, missing caveats, and accidental changes to
   un-commented sections. Requires `draft-scientific-paper`.
4. **Tone and concision** (`claude-sonnet-4-6`): make the prose direct, readable,
   and concise while reducing redundancy and preserving the scientific meaning.
   Requires `edit-scientific-prose`.

## How the agent workflow runs under Claude

`asta-revision run` is the complete launcher. It creates the run directory and all
step-1 outputs, runs an evidence preflight when a resolver is configured, then
invokes the agent workflow command, resolves required evidence requests, and
finalizes the Word/RIS outputs.

The coordinating step does no LLM work itself; it fans out to per-pass Claude
calls. So when `--agent-command` is not set (and no `ASTA_REVISION_AGENT_COMMAND`
is set), Step 2 defaults to calling the `asta-revision-agent` runner directly.

The runner spawns one headless Claude subagent per pass. When no subagent command
is set, it defaults to:

```text
claude -p --model {model} --add-dir {run_dir} --permission-mode acceptEdits
```

`{model}` is the per-pass model (opus for reasoning passes, sonnet for reviewer
passes). The subagent runs with `cwd=run_dir` so it discovers the run-local
`.claude/skills/` the launcher links in. `claude -p` prints its final message to
stdout; the runner captures that into the pass report. The subagent is instructed
not to edit files — it returns a report with an optional JSON payload of exact
markdown replacements, which the parent runner applies. Override the subagent
command with `--subagent-command` or `ASTA_REVISION_SUBAGENT_COMMAND`. If a skill
requires broader tools than `acceptEdits` allows, switch to
`--permission-mode bypassPermissions`.

Each subagent prompt names the required skills for that pass and requires the
report to mark skill-use checks such as `draft_scientific_paper_skill_used: true`;
if the nested session cannot access a required skill, the pass reports that as
false and the audit blocks finalization.

If revision implementation creates pending required evidence requests and no
resolver is configured, the runner stops before reviewer passes, avoiding
reviewer tokens on a draft already known to be blocked by missing evidence.

Finalization refuses to compile until
`agent_workflow/agent_workflow_audit.json` exists, names the same source DOCX
hash, hashes the exact `*.revised.md` being finalized, marks all four passes
complete, points to non-empty pass reports, and sets the required overall
readiness checks to true.

The final implementation is based only on the provided commented draft and its
comments. Do not rebuild from stale Markdown, TeX, or archived Word drafts unless
those files were generated from the same current `.docx`.

Launch the complete workflow:

```bash
export ASTA_REVISION_ASTA_COMMAND='asta-evidence-resolver --request {request_json} --output {output_json} --ris {output_ris}'

asta-revision run commented-draft.docx --output-stem manuscript_v4
```

## Install

```bash
uv tool install .
# or, for development:
uv pip install -e .
```

Requirements: `pandoc` and `claude` on PATH, plus `bip` (bipartite) if you use
the evidence resolver (see below).

## Evidence resolution with bipartite

Literature evidence is resolved through [bipartite](https://github.com/matsen/bipartite)
(`bip`): a git-backed JSONL bibliography ("nexus") with Semantic Scholar / Asta
search and DOI-based dedup. Install `bip` and configure
`~/.config/bip/config.yml` with `nexus_path`, `s2_api_key`, and `asta_api_key`.

For each evidence request, `asta-evidence-resolver`:

1. **Searches** the literature via bipartite.
2. **Dedups against the nexus** (reference-level): candidates whose DOI is
   already in the bibliography are dropped, so newly added citations are never
   redundant papers.
3. **Adds** the surviving new references to the nexus.
4. **Emits RIS** with complete records (type, title, author, year, plus DOI/venue
   when available).

> The exact `bip` subcommands are isolated in the `bip_*` wrapper functions of
> `asta_evidence_resolver.py`. Reconcile them with your installed `bip --help`
> and adjust in that one place if they differ; the resolver architecture
> (search → dedup vs nexus → add → RIS) is independent of the names.

### Three levels of redundancy guarding

1. **Reference-level** (bipartite, deterministic): never add a paper already in
   the nexus — DOI dedup above.
2. **Claim-level**: before adding a new cite-backed statement, the
   `asta_query_and_collation` pass checks `agent_workflow/cite_backed_statements.md`
   (the existing cite-backed sentences in the current revised markdown). If the
   same claim is already made-and-cited elsewhere, it reuses that citation instead
   of adding a redundant statement or a second citation. The pass must report
   `claim_redundancy_checked: true`.
3. **Prose-level**: the tone-and-concision pass plus `edit-scientific-prose`
   reduce repetitive phrasing.

`asta-revision finalize` resolves pending required requests before citation
generation, validates the returned RIS, writes `asta_reference_additions.json`,
and combines them with the embedded metadata as `citation_metadata.with_asta.ris`.
If evidence is required but unresolved, finalize fails and reports the pending
request IDs. The complete `run` launcher also uses the resolver for a preflight
before Step 2 starts.

## Word draft to EndNote-ready Word draft

```bash
asta-revision run commented-draft.docx \
  --output-stem manuscript_v4 \
  --asta-command 'asta-evidence-resolver --request {request_json} --output {output_json} --ris {output_ris}'
```

`asta-revision run` extracts text and style through Pandoc, extracts comments
directly from Word XML, extracts complete citation metadata from embedded EndNote
field records in the same source DOCX, runs an evidence preflight when a resolver
is configured, runs the configured agent workflow, resolves required evidence
requests, validates the required agent workflow audit, recompiles the revised
markdown, generates the paired RIS from the current recompiled reference list,
restores complete author/DOI fields from run-local metadata, converts numeric
citations to EndNote temporary citations, and runs sanity/sync checks.

`asta-revision start` and `asta-revision finalize` remain available as
lower-level debugging commands when a run needs manual inspection.

`launcher_profile.json` records per-step launcher timings, artifact sizes, and
approximate token counts. `agent_inputs/agent_input_manifest.json` names the
minimal recommended input files for each agent pass. Non-citation passes avoid
loading `citation_metadata.ris` by default. Evidence reviewers receive
`agent_inputs/citation_metadata_scoped.md`, a compact summary of only citation
records visible in Word-commented passages, plus compact evidence-response
summaries for newly resolved requests. Comment-scoped input files let revision,
rigor, and tone reviewers inspect only the Word-comment anchors unless they need
the full markdown. Rigor reviewers also receive `agent_workflow/scope_review.md`,
a whole-document paragraph-change summary; the asta pass receives
`agent_workflow/cite_backed_statements.md` for the claim-redundancy guard.

Citation handling is deterministic for a fixed run directory. The recompiled Word
document defines which references exist and their numbering/order. The run-local
`citation_metadata.ris` is built from embedded EndNote records in the current
source DOCX, including full author lists even when the visible Word bibliography
says `et al.`. During finalization, abbreviated visible references must match the
run-local metadata; otherwise the workflow fails rather than emitting truncated
authors.

`docx-reference-list-to-ris` remains available as a lower-level utility. It reads
the numbered reference list from a Word document and writes a matching RIS file;
pass `--metadata-ris --require-metadata-match` when the Word bibliography contains
`et al.` or otherwise abbreviated metadata.

`docx-numeric-to-endnote-temp` converts numeric superscript citations such as
`6,7` into EndNote temporary citations such as `{McCabe, 2012, Mutation of
A677...; McCabe, 2012, EZH2 inhibition...}`. When two distinct papers share the
same first author and year, the converter includes the title to make EndNote
matching unique. By default it removes the static reference list so EndNote can
regenerate the bibliography; use `--keep-references` for an inspection copy.

## EndNote import

1. Open the target library.
2. Choose `File > Import > File...`.
3. Select the generated `.ris` file.
4. Set `Import Option` to `Reference Manager (RIS)` or `RIS`.
5. Import the records.
6. Open the generated DOCX in Word and run EndNote's **Update Citations and
   Bibliography** command.

These tools emit EndNote temporary citations, not final EndNote fields. EndNote
performs the field conversion inside Word.

## Agentic revision workflow

The general workflow is documented in
[docs/agentic_revision_workflow.md](docs/agentic_revision_workflow.md). The core
rule is that the current commented `.docx` is the source of truth: revisions
should be made only in comment-scoped regions unless the review plan explicitly
justifies adjacent changes.

## Skills

Scientific-writing skills live in [skills/](skills/):
`draft-scientific-paper`, `edit-scientific-prose`, `asta-query-and-collation`,
`review-scientific-figures`, and `write-scientific-grant`.

The workflow wires `draft-scientific-paper`, `edit-scientific-prose`, and
`asta-query-and-collation` into the passes automatically: `asta-revision start`
symlinks the repo's `skills/` into each run directory's `.claude/skills/`, and the
runner launches each pass with `cwd=run_dir` so headless Claude discovers them as
project skills. No manual symlinking into a global skills directory is required.

## Repository name

If you cloned this from its previous name, rename your GitHub remote and local
directory to `asta-revision-workflow` to match the package.
