---
name: asta-query-and-collation
description: Resolve and collate literature evidence for manuscript claims that lack adjacent citation support, using a bipartite-backed reference store with paper-level and claim-level redundancy guards. Use when a revision pass needs to decide whether a modified or retained claim should get a new citation, query for evidence, and fold the results back into the run-local revision artifacts.
---

# Asta Query and Collation

Use this skill during the `asta_query_and_collation` pass of the Word revision
workflow. Its job is to make sure every retained claim that needs evidence is
either supported by an existing citation, supported by a newly resolved
reference, or explicitly softened — without introducing redundant citations.

Evidence resolution is backed by **bipartite** (`bip`): a git-backed JSONL
bibliography ("nexus") with Semantic Scholar / Asta search and DOI-based dedup.
Do not call the raw `asta` CLI directly; the configured resolver
(`asta-evidence-resolver`, invoked by the launcher) wraps `bip`.

## Inputs

- `agent_workflow/asta_requests.json` — the request ledger. Each required,
  pending entry is a claim that an earlier pass could not support from adjacent
  citations.
- `agent_workflow/cite_backed_statements.md` — the existing cite-backed
  sentences in the current revised markdown (regenerated each pass).
- `agent_workflow/asta/responses/*.json` and
  `agent_workflow/asta_reference_additions.json` — resolver outputs, when present.

## Workflow

1. **Claim-level redundancy guard (do this first).** For each pending required
   request, scan `cite_backed_statements.md` for a statement that already makes
   the same claim with a citation. If one exists:
   - Drop the request: set its `status` to `resolved_by_existing_citation` and
     record the matching anchor (the paragraph/sentence).
   - Reuse / cross-reference that existing citation rather than adding a new
     statement or a second citation for the same assertion.
   Keep a request only when the claim is genuinely new to the document.
2. **Resolve remaining requests.** For each request that is still required and
   pending, ensure the resolver has produced a complete RIS record. The resolver
   searches the literature, **dedups candidates against the nexus by DOI**
   (paper-level guard — never re-adds a paper already in the bibliography), adds
   the surviving new references to the nexus, and emits RIS.
3. **Collate.** Confirm resolved outputs are present in `agent_workflow/asta`,
   `agent_workflow/asta_reference_additions.json`, and the response JSON files so
   reviewer passes can see them.
4. **Confirm policy.** Verify that modified claims without adjacent evidence were
   resolved or intentionally softened; never leave an unsupported knowledge claim
   unresolved.

## Required report checks

Make these explicit in the final report (`true`/`false`):

- `modified_claims_citation_checked`
- `unsupported_claims_resolved`
- `asta_requests_collated`
- `claim_redundancy_checked` — set `true` once the step-1 redundancy scan has run.
- `source_docx_only`
- `draft_scientific_paper_skill_used`

## Principles

- Prefer reusing an existing citation over adding a new one. Two citations for
  the same assertion is redundancy, not rigor.
- A new citation is justified only when both guards pass: the *paper* is not
  already in the nexus (resolver/DOI dedup) and the *claim* is not already
  made-and-cited in the document (`cite_backed_statements.md`).
- When evidence cannot be found, soften or remove the claim rather than leaving
  it unsupported.
