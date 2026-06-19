#!/usr/bin/env python3
"""Resolve a pandoc revision Asta evidence request into JSON plus RIS.

The launcher calls this as:

    asta-evidence-resolver --request request.json --output response.json --ris output.ris

It intentionally remains a thin wrapper around the Asta CLI. The resolver fails
if Asta is unavailable or if no complete citation metadata can be emitted.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
from pathlib import Path
from typing import Any


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
        raise SystemExit("Asta request does not include request/query/claim text.")
    return "\n".join(parts)


def run_asta_find(query: str, output: Path, timeout: int) -> None:
    if shutil.which("asta") is None:
        raise SystemExit(
            "Asta CLI not found on PATH. Install it with `uv tool install --force "
            "git+https://github.com/allenai/asta-plugins.git@v0.17.1` or provide another resolver."
        )
    command = ["asta", "literature", "find", query, "-o", str(output), "--timeout", str(timeout)]
    print("+ " + " ".join(command))
    subprocess.run(command, check=True)


def run_asta_preflight(output: Path) -> None:
    if shutil.which("asta") is None:
        raise SystemExit(
            "Asta CLI not found on PATH. Install it with `uv tool install --force "
            "git+https://github.com/allenai/asta-plugins.git@v0.17.1`."
        )
    checks = [
        ["asta", "--version"],
        ["asta", "papers", "search", "H3K27ac", "--limit", "1"],
    ]
    transcripts: list[dict[str, Any]] = []
    for command in checks:
        print("+ " + " ".join(command))
        result = subprocess.run(command, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        transcripts.append({"command": command, "returncode": result.returncode, "output": result.stdout})
        if result.stdout:
            print(result.stdout)
        if result.returncode != 0:
            output.write_text(json.dumps({"status": "failed", "checks": transcripts}, indent=2) + "\n", encoding="utf-8")
            raise SystemExit(
                "Asta preflight failed. If authentication is required, use the URL printed by the Asta CLI above, "
                "then rerun the launcher."
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
    doi = clean_text(paper.get("doi") or paper.get("externalIds", {}).get("DOI") if isinstance(paper.get("externalIds"), dict) else "")
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


def load_results(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or not isinstance(data.get("results"), list):
        raise SystemExit(f"Asta output does not contain a results list: {path}")
    return data


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
        raise SystemExit("Asta search did not return any papers with title, author, and year.")
    path.write_text("\n\n".join(complete).strip() + "\n", encoding="utf-8")


def write_response(path: Path, request: dict[str, Any], query: str, asta_output: Path, papers: list[dict[str, Any]]) -> None:
    response = {
        "status": "resolved",
        "request": request,
        "query": query,
        "asta_output": str(asta_output),
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
        run_asta_preflight(output_path)
        ris_path.write_text(
            "TY  - JOUR\nTI  - Asta preflight placeholder\nAU  - Asta, Preflight\nPY  - 2026\nID  - asta-preflight\nER  -\n",
            encoding="utf-8",
        )
        print(f"Wrote Asta preflight response: {output_path}")
        return 0
    query = request_query(request)
    asta_output = output_path.with_suffix(".asta.json")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    ris_path.parent.mkdir(parents=True, exist_ok=True)

    run_asta_find(query, asta_output, args.timeout)
    data = load_results(asta_output)
    papers = select_papers(data["results"], args.limit)
    write_ris(ris_path, papers)
    write_response(output_path, request, query, asta_output, papers)
    print(f"Wrote Asta response: {output_path}")
    print(f"Wrote RIS: {ris_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
