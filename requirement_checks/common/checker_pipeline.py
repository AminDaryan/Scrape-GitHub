"""Shared single-pass checker pipeline.

Evaluates every paper via a per-paper callback, then reports and saves.
"""

import time

from common.fetch_and_parse_github_repo import parse_github_repo
from common.token_usage import print_token_usage_report


def run_checker_pipeline(
    *,
    papers,
    check_paper_fn,
    print_results_fn,
    save_results_fn,
    token_usage,
    deployment,
    description,
    github_token="",
    format_check_extra=None,
    finalize_fn=None,
):
    """Run the single-pass analysis pipeline.

    *check_paper_fn(paper)* returns a result dict.
    Optional *finalize_fn(results)* runs after printing (e.g. ground-truth).
    """
    print(
        f"Checking {len(papers)} repos for {description} "
        f"using LLM ({deployment})..."
    )
    if not github_token:
        print(
            "Tip: Set GITHUB_TOKEN in .env to avoid the "
            "60 req/hr GitHub rate limit.\n"
        )

    results = []

    for i, paper in enumerate(papers, 1):
        owner, repo_name = parse_github_repo(paper.get("repo", ""))
        label = f"{owner}/{repo_name}" if owner else paper.get("repo", "?")
        print(f"[{i:>2}/{len(papers)}] {label} ...", end=" ", flush=True)

        result = check_paper_fn(paper)
        results.append(result)

        icon = {
            "yes": "OK", "no": "NO", "skipped": "SKIP", "error": "ERR",
        }.get(result["status"], "?")
        extra = result.get("note") or ""
        suffix = (
            f"  {format_check_extra(result)}" if format_check_extra else ""
        )
        print(f"{icon}  {extra}{suffix}")
        time.sleep(0.3)

    # ── Reporting ─────────────────────────────────────────────────────────
    print_results_fn(results)

    if finalize_fn:
        finalize_fn(results)

    print_token_usage_report(token_usage, deployment)
    save_results_fn(results)
