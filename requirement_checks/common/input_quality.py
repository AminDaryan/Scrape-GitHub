"""Shared input data-quality validation — ONE entry point used before every check.

:func:`validate_input` checks a paper record's GitHub ``repo`` link (present /
parseable / duplicate / optionally live) plus its title.  The 5.1/5.2 checkers
call it via ``data.papers_source.load_papers``, so every run validates its input
as a preprocessing step.  An empty result means "nothing suspicious found", NOT a
guarantee of correctness.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))   # requirement_checks/
from common.github_helpers import parse_github_repo, is_github


def _liveness_flags(owner, name):
    """HEAD the repo to catch 404/moved links (network; opt-in)."""
    import requests
    try:
        resp = requests.head(f"https://github.com/{owner}/{name}", allow_redirects=True,
                             timeout=10, headers={"User-Agent": "input-quality/1.0"})
        if resp.status_code == 404:
            return [f"Repo not found on GitHub (404): {owner}/{name}"]
    except requests.RequestException:
        return [f"Repo liveness check could not complete (network): {owner}/{name}"]
    return []


def _repo_flags(paper, *, seen_repos=None, check_liveness=False):
    """Warnings for a GitHub-link paper's ``repo`` field."""
    repo = (paper.get("repo") or "").strip()
    if not repo:
        return ["Missing repo URL."]
    if not is_github(repo):
        return [f"Repo is not a GitHub URL (will be skipped): {repo}"]
    owner, name = parse_github_repo(repo)
    if not owner:
        return [f"Repo URL not parseable as github.com/owner/repo: {repo}"]
    flags = []
    if seen_repos is not None:
        key = f"{owner}/{name}".lower()
        if key in seen_repos:
            flags.append(f"Duplicate repo (also appears earlier in the list): {owner}/{name}")
        seen_repos.add(key)
    if check_liveness:
        flags += _liveness_flags(owner, name)
    return flags


def validate_input(paper, *, seen_repos=None, check_liveness=False):
    """All data-quality warnings for one record (title + GitHub repo checks)."""
    flags = []
    if not (paper.get("title") or "").strip():
        flags.append("Missing title.")
    if "repo" in paper:
        flags += _repo_flags(paper, seen_repos=seen_repos, check_liveness=check_liveness)
    else:
        flags.append("No repo link to validate.")
    return flags


def validate_papers(papers, *, check_liveness=False):
    """Return ``[(1-based index, paper, flags), ...]`` for papers with warnings."""
    seen, flagged = set(), []
    for i, paper in enumerate(papers, 1):
        flags = validate_input(paper, seen_repos=seen, check_liveness=check_liveness)
        if flags:
            flagged.append((i, paper, flags))
    return flagged


def log_report(papers, *, check_liveness=False):
    """Print a concise input data-quality report (the preprocessor step)."""
    flagged = validate_papers(papers, check_liveness=check_liveness)
    scope = "offline + liveness" if check_liveness else "offline"
    print(f"[input data quality] {len(papers)} paper(s); "
          f"{len(flagged)} with warnings ({scope}).")
    for i, paper, flags in flagged[:50]:
        title = (paper.get("title") or "?")[:60]
        print(f"  - #{i} {title!r}: " + "; ".join(flags))
    if len(flagged) > 50:
        print(f"  ... and {len(flagged) - 50} more")
    return flagged
