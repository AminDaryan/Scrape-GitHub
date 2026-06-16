"""Pre-flight input validation shared by the 5.1 / 5.2 checkers.

Runs as a preprocessor inside ``data.papers_source.load_papers()`` — so every
checker validates its (uploaded or vendored) papers BEFORE evaluating them and
prints a short data-quality report (the UI surfaces it in the run log).  It
catches input problems that would otherwise only surface mid-run or silently
skew results: missing title/repo, non-GitHub or unparseable repo links,
duplicate repos, and — when ``CHECK_REPO_LIVENESS=1`` — dead/moved repo URLs.

These papers carry GitHub links, so the checks are repo-focused.  The Beall's
check validates Semantic Scholar *venue* metadata separately, in
``bealls_list_check/data_quality.py`` (which annotates its Excel output).
"""

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))   # requirement_checks/
from common.github_helpers import parse_github_repo, is_github


def _liveness_flags(owner, name):
    """HEAD the repo to catch 404/moved links (network; opt-in)."""
    import requests
    try:
        resp = requests.head(f"https://github.com/{owner}/{name}",
                             allow_redirects=True, timeout=10,
                             headers={"User-Agent": "input-quality/1.0"})
        if resp.status_code == 404:
            return [f"Repo not found on GitHub (404): {owner}/{name}"]
    except requests.RequestException:
        return [f"Repo liveness check could not complete (network): {owner}/{name}"]
    return []


def paper_flags(paper, *, seen_repos=None, check_liveness=False):
    """Data-quality warnings for one input paper (empty list = looks fine)."""
    flags = []
    if not (paper.get("title") or "").strip():
        flags.append("Missing title.")

    repo = (paper.get("repo") or "").strip()
    if not repo:
        flags.append("Missing repo URL.")
        return flags
    if not is_github(repo):
        flags.append(f"Repo is not a GitHub URL (will be skipped): {repo}")
        return flags

    owner, name = parse_github_repo(repo)
    if not owner:
        flags.append(f"Repo URL not parseable as github.com/owner/repo: {repo}")
        return flags

    if seen_repos is not None:
        key = f"{owner}/{name}".lower()
        if key in seen_repos:
            flags.append(f"Duplicate repo (also appears earlier in the list): {owner}/{name}")
        seen_repos.add(key)
    if check_liveness:
        flags += _liveness_flags(owner, name)
    return flags


def validate_papers(papers, *, check_liveness=False):
    """Return ``[(1-based index, paper, flags), ...]`` for papers with warnings."""
    seen, flagged = set(), []
    for i, paper in enumerate(papers, 1):
        flags = paper_flags(paper, seen_repos=seen, check_liveness=check_liveness)
        if flags:
            flagged.append((i, paper, flags))
    return flagged


def log_report(papers, *, check_liveness=False):
    """Print a concise input data-quality report (the preprocessor step)."""
    flagged = validate_papers(papers, check_liveness=check_liveness)
    scope = "offline + liveness" if check_liveness else "offline only"
    print(f"[input data quality] {len(papers)} paper(s); "
          f"{len(flagged)} with warnings ({scope}).")
    for i, paper, flags in flagged[:50]:
        title = (paper.get("title") or "?")[:60]
        print(f"  - #{i} {title!r}: " + "; ".join(flags))
    if len(flagged) > 50:
        print(f"  ... and {len(flagged) - 50} more")
    return flagged
