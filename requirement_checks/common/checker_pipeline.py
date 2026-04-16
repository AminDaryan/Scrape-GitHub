"""Shared single-pass checker pipeline.

Evaluates every paper via a per-paper callback, then reports and saves.
Supports both single-model and multi-model runs.
"""

import time

from common.fetch_and_parse_github_repo import parse_github_repo
from common.token_usage import TokenUsageTracker, print_token_usage_report


def _run_single_pass(
    *,
    papers,
    check_paper_fn,
    deployment,
    description,
    github_token="",
    format_check_extra=None,
):
    """Run check_paper_fn on every paper. Returns list of result dicts."""
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

    return results


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
    """Run the single-pass analysis pipeline (single model).

    *check_paper_fn(paper)* returns a result dict.
    Optional *finalize_fn(results)* runs after printing (e.g. ground-truth).
    """
    results = _run_single_pass(
        papers=papers,
        check_paper_fn=check_paper_fn,
        deployment=deployment,
        description=description,
        github_token=github_token,
        format_check_extra=format_check_extra,
    )

    print_results_fn(results)
    if finalize_fn:
        finalize_fn(results)
    print_token_usage_report(token_usage, deployment)
    save_results_fn(results)


def run_multi_model_pipeline(
    *,
    papers,
    deployments,
    make_check_paper_fn,
    print_results_fn,
    save_results_fn,
    description,
    github_token="",
    format_check_extra=None,
    finalize_fn=None,
):
    """Run the analysis pipeline across multiple LLM deployments.

    Parameters
    ----------
    make_check_paper_fn : callable(deployment, token_usage) -> check_paper_fn
        Factory that creates a per-paper checker bound to a specific model.
    save_results_fn : callable(all_model_results)
        Receives ``{deployment: {"results": [...], "token_usage": tracker}}``.
    """
    all_model_results = {}

    for idx, deployment in enumerate(deployments):
        if len(deployments) > 1:
            print(f"\n{'#' * 78}")
            print(f"# MODEL {idx + 1}/{len(deployments)}: {deployment}")
            print(f"{'#' * 78}\n")

        token_usage = TokenUsageTracker()
        check_paper_fn = make_check_paper_fn(deployment, token_usage)

        results = _run_single_pass(
            papers=papers,
            check_paper_fn=check_paper_fn,
            deployment=deployment,
            description=description,
            github_token=github_token,
            format_check_extra=format_check_extra,
        )

        print_results_fn(results)
        if finalize_fn:
            finalize_fn(results)
        print_token_usage_report(token_usage, deployment)

        all_model_results[deployment] = {
            "results": results,
            "token_usage": token_usage,
        }

    save_results_fn(all_model_results)
