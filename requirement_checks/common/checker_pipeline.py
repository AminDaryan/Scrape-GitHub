"""Runs the checking process for all papers and saves the results.

For each paper in the list, this module calls a checker function (e.g. one
that checks for inline comments or installation instructions).  It loops
through every paper, prints progress to the console, collects the results,
and writes them to an Excel file.

Entry point:
  - run_pipeline()   — checks all papers with one or more LLM deployments
                       and saves per-model results (plus a comparison sheet
                       when >1 model).  Accepts deployments as a single
                       string or a list of strings.
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
    token_usage=None,
):
    """Loop through every paper, call check_paper_fn on each, print progress.

    For each paper:
      1. Extract the GitHub owner/repo from the paper's URL.
      2. Call check_paper_fn(paper), which sends repo content to the LLM and
         returns a result dict with at least a "status" key (yes/no/skipped/error).
      3. Print a one-line status to the console (e.g. "[ 3/20] owner/repo ... OK").

    Returns the list of all result dicts (one per paper).

    Parameters
    ----------
    papers : list[dict]
        The paper entries from papers_from_database.py.
    check_paper_fn : callable(paper) -> dict
        The checker function that evaluates a single paper's repo.
    deployment : str
        Name of the LLM deployment (e.g. "gpt-5-2025-08-07"), shown in logs.
    description : str
        What is being checked (e.g. "inline comments"), shown in the header.
    github_token : str
        Optional GitHub API token; prints a rate-limit warning if empty.
    format_check_extra : callable(result) -> str or None
        Optional function to append extra info to each console line.
    token_usage : TokenUsageTracker or None
        When provided, the number of tokens consumed for each paper is stored
        in ``result["tokens_used"]`` so it can be written to Excel.
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

        tokens_before = token_usage.total_tokens if token_usage is not None else 0
        result = check_paper_fn(paper)
        result["tokens_used"] = (
            token_usage.total_tokens - tokens_before
            if token_usage is not None else None
        )
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


def run_pipeline(
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
    """Check all papers with one or more LLM deployments and save results.

    deployments can be a single string (e.g. "gpt-5") or a list of strings
    (e.g. ["gpt-5", "gpt-5-mini"]).  A single string is automatically
    wrapped into a one-element list.

    For each deployment in the list:
      1. Create a fresh TokenUsageTracker for that model.
      2. Call make_check_paper_fn(deployment, token_usage) to get a checker
         function bound to that specific model.
      3. Call _run_single_pass() to check every paper with that checker.
      4. Print results and token usage for that model.

    After all deployments finish:
      5. Call save_results_fn(all_model_results) with a dict like:
         {"gpt-5": {"results": [...], "token_usage": tracker}, "gpt-5-mini": ...}
         This writes per-model sheets plus a Model Comparison sheet (when >1 model).

    Parameters
    ----------
    papers : list[dict]
        The paper entries to check.
    deployments : str or list[str]
        One deployment name or a list of deployment names.
    make_check_paper_fn : callable(deployment, token_usage) -> callable
        Factory: given a deployment name and a tracker, returns a
        check_paper_fn(paper) -> dict for that model.
    print_results_fn : callable(results)
        Prints a human-readable summary after each model finishes.
    save_results_fn : callable(all_model_results)
        Writes the combined results to an Excel file with comparison.
    description : str
        What is being checked (e.g. "usage examples").
    github_token : str
        Optional GitHub API token.
    format_check_extra : callable or None
        Optional extra info formatter for console output.
    finalize_fn : callable(results) or None
        Optional post-processing step run after each model.
    """
    if isinstance(deployments, str):
        deployments = [deployments]

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
            token_usage=token_usage,
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
