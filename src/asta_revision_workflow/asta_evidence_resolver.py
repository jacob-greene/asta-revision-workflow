#!/usr/bin/env python3
"""Resolve a revision evidence request into JSON plus RIS, backed by bipartite.

The launcher calls this as:

    asta-evidence-resolver --request request.json --output response.json --ris output.ris

Evidence is resolved through the bipartite CLI (``bip``): a git-backed JSONL
bibliography ("nexus") with Semantic Scholar / Asta search and DOI-based dedup.
Per request the resolver searches the literature, drops candidates whose DOI is
already in the nexus (reference-level dedup), adds the surviving new references to
the nexus, and emits complete RIS. It fails if ``bip`` is unavailable or if no
complete citation metadata can be emitted.

NOTE: the exact ``bip`` subcommands/flags below come from the bipartite README
and are isolated in the ``bip_*`` wrapper functions. Reconcile them against the
installed ``bip --help`` and adjust in one place if they differ; the resolver
architecture (search -> dedup vs nexus -> add -> RIS) is independent of the names.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
from pathlib import Path
from typing import Any

# Isolated bipartite subcommands. Reconcile with `bip --help` (see module docstring).
BIP_SEARCH_ARGS = ("search",)           # bip search "<query>" --json
BIP_LIST_ARGS = ("list",)               # bip list --json  (enumerate nexus entries for DOI dedup)
BIP_ADD_ARGS = ("s2", "add")            # bip s2 add DOI:<doi>


def request_query(request: dict[str, Any]) -> str:
    parts: list[str] = []
    for key in ("request", "query", "claim", "needed_evidence", "purpose"):
        value = request.get(key)
        if isinstance(value, str) and value.strip():
            parts.append(value.strip())
    needed_claims = request.get("needed_claims")
    if isinstance(needed_claims, list):
        parts.extend(str(item).strip() for item in needed_claims if str(item).strip())
    if not parts:
        raise SystemExit("Evidence request does not include request/query/claim text.")
    return "\n".join(parts)


def require_bip() -> None:
    if shutil.which("bip") is None:
        raise SystemExit(
            "bipartite CLI (`bip`) not found on PATH. Install it (https://github.com/matsen/bipartite) "
            "and configure ~/.config/bip/config.yml with nexus_path, s2_api_key, and asta_api_key, "
            "or provide another resolver via --asta-command / ASTA_REVISION_ASTA_COMMAND."
        )


def run_bip(args: tuple[str, ...] | list[str], *, check: bool = True) -> subprocess.CompletedProcess[str]:
    command = ["bip", *args]
    print("+ " + " ".join(command))
    return subprocess.run(command, check=check, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)


def normalize_doi(value: Any) -> str:
    doi = clean_text(value).lower()
    doi = re.sub(r"^(?:https?://(?:dx\.)?doi\.org/|doi:)", "", doi)
    return doi


def paper_doi(paper: dict[str, Any]) -> str:
    external = paper.get("externalIds")
    raw = paper.get("doi") or (external.get("DOI") if isinstance(external, dict) else "")
    return normalize_doi(raw)


def normalize_bip_paper(item: dict[str, Any]) -> dict[str, Any]:
    """Map a bipartite search result onto the S2-shaped dict the RIS helpers expect.

    Tolerant of field-name variation between bipartite versions; reconcile with
    actual `bip search --json` output if a field is missing.
    """
    if not isinstance(item, dict):
        return {}
    authors = item.get("authors")
    if isinstance(authors, list):
        norm_authors = [a if isinstance(a, (dict, str)) else str(a) for a in authors]
    elif isinstance(authors, str):
        norm_authors = [a.strip() for a in re.split(r";|\band\b", authors) if a.strip()]
    else:
        norm_authors = []
    return {
        "title": item.get("title"),
        "year": item.get("year") or item.get("publicationYear"),
        "authors": norm_authors,
        "venue": item.get("venue") or item.get("journal") or item.get("publicationVenue"),
        "doi": item.get("doi") or "",
        "externalIds": item.get("externalIds") if isinstance(item.get("externalIds"), dict) else {},
        "corpusId": item.get("corpusId") or item.get("corpus_id"),
        "url": item.get("url"),
        "abstract": item.get("abstract"),
        "relevanceScore": item.get("relevanceScore") or item.get("score") or 0,
        "relevanceJudgement": item.get("relevanceJudgement"),
    }


def parse_bip_papers(stdout: str) -> list[dict[str, Any]]:
    """Parse `bip` JSON stdout into a list of normalized paper dicts.

    Accepts either a top-level list, a {"results": [...]} / {"papers": [...]}
    object, or JSON-lines output.
    """
    stdout = stdout.strip()
    if not stdout:
        return []
    items: list[Any] = []
    try:
        data = json.loads(stdout)
        if isinstance(data, list):
            items = data
        elif isinstance(data, dict):
            for key in ("results", "papers", "entries", "items"):
                if isinstance(data.get(key), list):
                    items = data[key]
                    break
            else:
                items = [data]
    except json.JSONDecodeError:
        for line in stdout.splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                items.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    return [normalize_bip_paper(item) for item in items if isinstance(item, dict)]


def bip_search(query: str, timeout: int) -> list[dict[str, Any]]:
    """Search the literature via bipartite's S2/Asta backend."""
    result = run_bip([*BIP_SEARCH_ARGS, query, "--json", "--timeout", str(timeout)])
    return parse_bip_papers(result.stdout)


def bip_nexus_dois() -> set[str]:
    """Return the set of DOIs already in the nexus, for reference-level dedup."""
    result = run_bip([*BIP_LIST_ARGS, "--json"], check=False)
    if result.returncode != 0:
        return set()
    dois: set[str] = set()
    for paper in parse_bip_papers(result.stdout):
        doi = paper_doi(paper)
        if doi:
            dois.add(doi)
    return dois


def bip_add(doi: str) -> None:
    """Add a reference to the nexus by DOI (bipartite dedups on its side)."""
    if not doi:
        return
    run_bip([*BIP_ADD_ARGS, f"DOI:{doi}"], check=False)


def run_bip_preflight(output: Path) -> None:
    require_bip()
    transcripts: list[dict[str, Any]] = []
    result = run_bip(["--help"], check=False)
    transcripts.append({"command": ["bip", "--help"], "returncode": result.returncode, "output": result.stdout})
    if result.returncode != 0:
        output.write_text(json.dumps({"status": "failed", "checks": transcripts}, indent=2) + "\n", encoding="utf-8")
        raise SystemExit(
            "bipartite preflight failed. Ensure `bip` is installed and ~/.config/bip/config.yml is configured "
            "(nexus_path, s2_api_key, asta_api_key), then rerun the launcher."
        )
    output.write_text(json.dumps({"status": "ok", "checks": transcripts}, indent=2) + "\n", encoding="utf-8")


def clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def ris_escape(value: Any) -> str:
    text = clean_text(value)
    return text.replace("\n", " ")


def paper_id(paper: dict[str, Any], index: int) -> str:
    title = clean_text(paper.get("title"))
    year = clean_text(paper.get("year"))
    first_author = ""
    authors = paper.get("authors")
    if isinstance(authors, list) and authors:
        first = authors[0]
        if isinstance(first, dict):
            first_author = clean_text(first.get("name")).split()[-1]
        else:
            first_author = clean_text(first).split()[-1]
    slug = re.sub(r"[^A-Za-z0-9]+", "", f"{first_author}{year}{title[:30]}")
    return slug or f"astaEvidence{index}"


def paper_to_ris(paper: dict[str, Any], index: int) -> str | None:
    title = clean_text(paper.get("title"))
    year = clean_text(paper.get("year"))
    authors = paper.get("authors")
    if not title or not year or not isinstance(authors, list) or not authors:
        return None
    lines = ["TY  - JOUR", f"TI  - {ris_escape(title)}"]
    for author in authors:
        if isinstance(author, dict):
            name = clean_text(author.get("name"))
        else:
            name = clean_text(author)
        if name:
            lines.append(f"AU  - {ris_escape(name)}")
    lines.append(f"PY  - {ris_escape(year)}")
    venue = clean_text(paper.get("venue") or paper.get("journal"))
    if venue:
        lines.append(f"JO  - {ris_escape(venue)}")
    doi = paper_doi(paper)
    if doi:
        lines.append(f"DO  - {ris_escape(doi)}")
    corpus_id = clean_text(paper.get("corpusId"))
    if corpus_id:
        lines.append(f"AN  - {ris_escape(corpus_id)}")
    url = clean_text(paper.get("url"))
    if url:
        lines.append(f"UR  - {ris_escape(url)}")
    abstract = clean_text(paper.get("abstract"))
    if abstract:
        lines.append(f"AB  - {ris_escape(abstract)}")
    lines.append(f"ID  - {paper_id(paper, index)}")
    lines.append("ER  -")
    return "\n".join(lines)


def select_papers(results: list[dict[str, Any]], limit: int) -> list[dict[str, Any]]:
    scored = [item for item in results if isinstance(item, dict)]
    scored.sort(key=lambda item: float(item.get("relevanceScore") or 0), reverse=True)
    selected: list[dict[str, Any]] = []
    for paper in scored:
        if paper_to_ris(paper, len(selected) + 1) is None:
            continue
        selected.append(paper)
        if len(selected) >= limit:
            break
    return selected


def write_ris(path: Path, papers: list[dict[str, Any]]) -> None:
    records = [paper_to_ris(paper, index) for index, paper in enumerate(papers, start=1)]
    complete = [record for record in records if record]
    if not complete:
        raise SystemExit("Literature search did not return any papers with title, author, and year.")
    path.write_text("\n\n".join(complete).strip() + "\n", encoding="utf-8")


def write_response(
    path: Path,
    request: dict[str, Any],
    query: str,
    raw_output: Path,
    papers: list[dict[str, Any]],
    *,
    candidate_count: int = 0,
    redundant_filtered: int = 0,
) -> None:
    response = {
        "status": "resolved",
        "backend": "bipartite",
        "request": request,
        "query": query,
        "raw_output": str(raw_output),
        "candidate_count": candidate_count,
        "redundant_filtered_by_nexus_doi": redundant_filtered,
        "added_to_nexus": [paper_doi(paper) for paper in papers if paper_doi(paper)],
        "selected_papers": [
            {
                "title": paper.get("title"),
                "year": paper.get("year"),
                "venue": paper.get("venue"),
                "corpusId": paper.get("corpusId"),
                "url": paper.get("url"),
                "relevanceScore": paper.get("relevanceScore"),
                "relevanceSummary": (
                    paper.get("relevanceJudgement", {}).get("relevanceSummary")
                    if isinstance(paper.get("relevanceJudgement"), dict)
                    else None
                ),
            }
            for paper in papers
        ],
    }
    path.write_text(json.dumps(response, indent=2) + "\n", encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--request", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--ris", required=True)
    parser.add_argument("--limit", type=int, default=3)
    parser.add_argument("--timeout", type=int, default=300)
    args = parser.parse_args(argv)

    request_path = Path(args.request).resolve()
    output_path = Path(args.output).resolve()
    ris_path = Path(args.ris).resolve()
    request = json.loads(request_path.read_text(encoding="utf-8"))
    if request.get("preflight"):
        output_path.parent.mkdir(parents=True, exist_ok=True)
        ris_path.parent.mkdir(parents=True, exist_ok=True)
        run_bip_preflight(output_path)
        ris_path.write_text(
            "TY  - JOUR\nTI  - bipartite preflight placeholder\nAU  - Bipartite, Preflight\nPY  - 2026\nID  - bip-preflight\nER  -\n",
            encoding="utf-8",
        )
        print(f"Wrote bipartite preflight response: {output_path}")
        return 0

    require_bip()
    query = request_query(request)
    raw_output = output_path.with_suffix(".bip.json")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    ris_path.parent.mkdir(parents=True, exist_ok=True)

    # 1. Search the literature via bipartite.
    candidates = bip_search(query, args.timeout)
    raw_output.write_text(json.dumps({"query": query, "results": candidates}, indent=2) + "\n", encoding="utf-8")

    # 2. Reference-level dedup: drop candidates whose DOI is already in the nexus.
    nexus_dois = bip_nexus_dois()
    new_candidates = [paper for paper in candidates if paper_doi(paper) not in nexus_dois or not paper_doi(paper)]
    redundant_filtered = len(candidates) - len(new_candidates)

    # 3. Select complete records, then add the survivors to the nexus.
    papers = select_papers(new_candidates, args.limit)
    for paper in papers:
        bip_add(paper_doi(paper))

    # 4. Emit RIS + a response for reviewer visibility.
    write_ris(ris_path, papers)
    write_response(
        output_path,
        request,
        query,
        raw_output,
        papers,
        candidate_count=len(candidates),
        redundant_filtered=redundant_filtered,
    )
    print(f"Wrote evidence response: {output_path}")
    print(f"Wrote RIS: {ris_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
