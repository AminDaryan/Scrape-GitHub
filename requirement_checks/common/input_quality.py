"""Shared input data-quality validation — ONE entry point used before EVERY check.

:func:`validate_input` is field-aware, so the 5.1/5.2 checkers (via
``data.papers_source.load_papers``) and the Beall's check (via
``bealls_list_check.classify_corpus``) run the *same* function.  It dispatches on
what the record actually contains:

  * a GitHub ``repo`` link  -> repo checks (present / parseable / duplicate /
    optionally live);
  * Semantic Scholar venue fields (``publicationVenue`` / ``venue`` / DOI / ISSN)
    -> venue checks (delegated to ``bealls_list_check.data_quality``: missing
    venue, malformed DOI/ISSN, encoding artifacts, "merged venues", and an
    optional Crossref cross-check).

A record can have both (both run), or neither (flagged as such).  An empty result
means "nothing suspicious found", NOT a guarantee of correctness.
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


def _venue_flags(paper, *, use_crossref=False):
    """Venue checks, delegated to bealls_list_check.data_quality (lazy import)."""
    bealls = Path(__file__).resolve().parent.parent / "bealls_list_check"
    if str(bealls) not in sys.path:
        sys.path.insert(0, str(bealls))
    try:
        import data_quality
    except Exception:
        return []
    flags = data_quality.offline_flags(paper)
    if use_crossref:
        flags += data_quality.crossref_flags(paper)
    return flags


def _has_venue_fields(paper):
    return bool(paper.get("publicationVenue")) or bool(paper.get("venue")) \
        or bool((paper.get("externalIds") or {}).get("DOI"))


def validate_input(paper, *, seen_repos=None, check_liveness=False, use_crossref=False):
    """All data-quality warnings for one record, dispatched on the fields present."""
    flags = []
    if not (paper.get("title") or "").strip():
        flags.append("Missing title.")
    has_repo = "repo" in paper
    has_venue = _has_venue_fields(paper)
    if has_repo:
        flags += _repo_flags(paper, seen_repos=seen_repos, check_liveness=check_liveness)
    if has_venue:
        flags += _venue_flags(paper, use_crossref=use_crossref)
    if not has_repo and not has_venue:
        flags.append("No repo link or venue metadata to validate.")
    return flags


def validate_papers(papers, *, check_liveness=False, use_crossref=False):
    """Return ``[(1-based index, paper, flags), ...]`` for papers with warnings."""
    seen, flagged = set(), []
    for i, paper in enumerate(papers, 1):
        flags = validate_input(paper, seen_repos=seen, check_liveness=check_liveness,
                               use_crossref=use_crossref)
        if flags:
            flagged.append((i, paper, flags))
    return flagged


def log_report(papers, *, check_liveness=False, use_crossref=False):
    """Print a concise input data-quality report (the preprocessor step)."""
    flagged = validate_papers(papers, check_liveness=check_liveness, use_crossref=use_crossref)
    scope = "offline + liveness" if check_liveness else "offline"
    print(f"[input data quality] {len(papers)} paper(s); "
          f"{len(flagged)} with warnings ({scope}).")
    for i, paper, flags in flagged[:50]:
        title = (paper.get("title") or "?")[:60]
        print(f"  - #{i} {title!r}: " + "; ".join(flags))
    if len(flagged) > 50:
        print(f"  ... and {len(flagged) - 50} more")
    return flagged
