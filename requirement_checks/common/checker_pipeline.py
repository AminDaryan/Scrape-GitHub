"""Shared two-pass checker pipeline.

Both checkers follow the same workflow: evaluate every paper (pass 1),
auto-heal low-confidence results (pass 2), then report and save.
This module extracts that loop so each checker only supplies callbacks.
"""

import time

from common.fetch_and_parse_github_repo import parse_github_repo
from common.confidence_reporting import print_confidence_report
from common.token_usage import print_token_usage_report


def run_checker_pipeline(
    *,
    papers,
    check_paper_fn,
    diagnose_fn,
    heal_fn,
    print_results_fn,
    save_results_fn,
    token_usage,
    deployment,
    description,
    github_token="",
    format_check_extra=None,
    format_heal_extra=None,
    finalize_fn=None,
):
    """Run the standard two-pass analysis pipeline.

    *check_paper_fn(paper)* returns ``(result_dict, content_str)``.
    *heal_fn(result, content, owner, repo)* returns a healed result.
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
    repo_contents = {}

    # ── Pass 1: evaluate every paper ──────────────────────────────────────
    for i, paper in enumerate(papers, 1):
        owner, repo_name = parse_github_repo(paper.get("repo", ""))
        label = f"{owner}/{repo_name}" if owner else paper.get("repo", "?")
        print(f"[{i:>2}/{len(papers)}] {label} ...", end=" ", flush=True)

        result, content = check_paper_fn(paper)
        results.append(result)
        repo_contents[paper["title"]] = content

        icon = {
            "yes": "OK", "no": "NO", "skipped": "SKIP", "error": "ERR",
        }.get(result["status"], "?")
        extra = result.get("confidence") or result.get("note") or ""
        suffix = (
            f"  {format_check_extra(result)}" if format_check_extra else ""
        )
        print(f"{icon}  {extra}{suffix}")
        time.sleep(0.3)

    # ── Confidence report ─────────────────────────────────────────────────
    print_confidence_report(results, repo_contents, diagnose_fn)

    # ── Pass 2: self-heal medium/low confidence results ───────────────────
    non_high = [
        i for i, r in enumerate(results)
        if r.get("confidence") in ("medium", "low")
    ]
    if non_high:
        print(
            f"\n  Auto-healing {len(non_high)} medium/low "
            f"confidence result(s)...\n"
        )
        for idx in non_high:
            r = results[idx]
            owner, repo = parse_github_repo(r.get("repo", ""))
            if not owner:
                continue
            title_short = r["title"][:55] + (
                "…" if len(r["title"]) > 55 else ""
            )
            print(f"  healing: {title_short}", end=" ... ", flush=True)

            healed = heal_fn(
                r, repo_contents.get(r["title"], ""), owner, repo,
            )
            results[idx] = healed

            heal_extra = (
                f"  {format_heal_extra(r, healed)}"
                if format_heal_extra else ""
            )
            print(
                f"conf: {r.get('confidence')} → "
                f"{healed.get('confidence')}{heal_extra}"
            )
            time.sleep(0.5)

        still_non_high = [
            r for r in results
            if r.get("confidence") in ("medium", "low")
        ]
        if still_non_high:
            print(
                f"\n  {len(still_non_high)} result(s) still not "
                f"high confidence — manual review needed."
            )
        else:
            print("\n  All results are now high confidence.")

    # ── Reporting ─────────────────────────────────────────────────────────
    print_results_fn(results)

    if finalize_fn:
        finalize_fn(results)

    print_token_usage_report(token_usage, deployment)
    save_results_fn(results)
