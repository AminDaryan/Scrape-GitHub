"""GitHub helpers — URL parsing, API access, and bulk content fetching.

This module is the single place where the project talks to GitHub.  It is
organised in three layers, from low-level to high-level:

  1. URL handling
       parse_github_repo(url)  ->  (owner, repo)
       is_github(url)          ->  bool

  2. Raw GitHub REST API access
       github_get(path)                    ->  parsed JSON or None
       github_get_with_link_header(path)   ->  (json, link_header) for pagination
       list_all_repo_files(o, r)           ->  list of file blobs
       fetch_file_content(o, r, p)         ->  decoded UTF-8 text or None

  3. Higher-level helpers used by checker scripts
       fetch_paths_with_char_budget(...)  ->  fetches many files into a
            single LLM-friendly payload that respects a max-character cap
       path_priority_with_readme_first(...)  ->  sort key for ordering
            files so that root READMEs come first, etc.
       first_readme_path(...)  ->  finds the root README in a file listing.

Authentication
--------------
Set ``GITHUB_TOKEN`` in ``.env`` to lift the unauthenticated rate limit
(60 requests/hour) up to 5,000 requests/hour.  Without a token, large
multi-paper runs will hit the limit.
"""

import base64
import json
import os
import re
import time
import urllib.parse
from typing import Any, Callable, List, Mapping, Optional, Sequence, Tuple
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from dotenv import load_dotenv

load_dotenv()

GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")


# =============================================================================
# 1. URL HANDLING
# =============================================================================

def parse_github_repo(url):
    """Extract ``(owner, repo)`` from a GitHub URL.

    Handles common variations:
      - https://github.com/owner/repo
      - https://github.com/owner/repo/
      - https://github.com/owner/repo.git
      - https://github.com/owner/repo/blob/main/path/to/file

    Returns ``(None, None)`` if the URL is not a recognisable GitHub repo URL
    (e.g. HuggingFace links, plain project websites).  Callers should treat
    a None result as "not a GitHub repo, skip it".
    """
    match = re.search(r"github\.com/([^/]+)/([^/?.#]+)", url)
    if match:
        owner = match.group(1)
        repo  = match.group(2).rstrip("/")
        if repo.endswith(".git"):
            repo = repo[:-4]
        return owner, repo
    return None, None


def is_github(url):
    """Return True if *url* points to a GitHub repository.

    A simple substring check — does not validate that the URL actually
    resolves or that the repo exists.  Used for fast filtering before
    deciding whether to spend an API call on a paper.
    """
    return "github.com" in url


# =============================================================================
# 2. RAW GITHUB REST API ACCESS
# =============================================================================

def github_get(path):
    """GET a GitHub REST API endpoint and return parsed JSON.

    Returns:
        dict | list | None: Parsed JSON on success.  None on 4xx (other
        than 403), network failure, or timeout.

    Raises:
        RuntimeError: When the GitHub rate limit is exhausted (HTTP 403).
            This is raised rather than swallowed because it usually means
            the entire script run cannot continue — adding a token or
            waiting an hour is required.
    """
    url     = f"https://api.github.com/{path.lstrip('/')}"
    headers = {"Accept": "application/vnd.github.v3+json"}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    req = Request(url, headers=headers)
    try:
        with urlopen(req, timeout=10) as r:
            return json.loads(r.read().decode())
    except HTTPError as e:
        if e.code == 403:
            raise RuntimeError(
                "GitHub API rate limit hit. Set GITHUB_TOKEN in .env or wait an hour."
            )
        return None
    except URLError:
        return None


def github_get_with_link_header(path):
    """Like ``github_get`` but also returns the response's ``Link`` header.

    GitHub paginates list endpoints (commits, issues, releases, …) and
    exposes the cursor only via the ``Link`` HTTP header.  The standard
    way to count total items without walking every page is to ask for
    ``per_page=1`` and then parse the ``rel="last"`` URL — its ``page=N``
    parameter equals the total count.

    Returns ``(parsed_json, link_str)``:
      * ``parsed_json`` — same as ``github_get`` (dict | list | None).
      * ``link_str``    — the raw ``Link`` header, or ``""`` when the
        response was a single page.
    """
    url     = f"https://api.github.com/{path.lstrip('/')}"
    headers = {"Accept": "application/vnd.github.v3+json"}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    req = Request(url, headers=headers)
    try:
        with urlopen(req, timeout=10) as r:
            data = json.loads(r.read().decode())
            return data, r.headers.get("Link", "") or ""
    except HTTPError as e:
        if e.code == 403:
            raise RuntimeError(
                "GitHub API rate limit hit. Set GITHUB_TOKEN in .env or wait an hour."
            )
        return None, ""
    except URLError:
        return None, ""


def list_all_repo_files(owner, repo):
    """Return every file blob in the repo using the recursive Git Trees API.

    This is a single API call regardless of repo size — the entire repo tree
    is returned at once, which is much cheaper than walking directories with
    the Contents API (one call per directory).

    Returns:
        list[dict]: Each dict has keys ``path``, ``type`` (always "blob" in
        the result), ``sha``, and ``size``.  Returns ``[]`` if the repo is
        empty, the response is malformed, or the API call fails.

    Note:
        Trees over 100,000 entries are truncated by GitHub.  For most
        academic repos this limit is irrelevant.
    """
    data = github_get(f"repos/{owner}/{repo}/git/trees/HEAD?recursive=1")
    if not data or "tree" not in data:
        return []
    return [item for item in data["tree"] if item.get("type") == "blob"]


def fetch_file_content(owner, repo, file_path):
    """Download one file from GitHub and return its UTF-8 text.

    Uses the Contents API which returns base64-encoded file data.
    File paths are URL-encoded so that paths containing spaces, ``#``,
    or other special characters work correctly.

    Returns:
        str | None: The decoded file content, or None if the file is
        missing, binary (decoding fails), or the API call fails.
        Errors during UTF-8 decoding are silently ignored (``errors="ignore"``)
        so binary blobs return as empty/garbled strings rather than
        raising.
    """
    encoded_path = urllib.parse.quote(file_path, safe="/")
    data = github_get(f"repos/{owner}/{repo}/contents/{encoded_path}")
    if not data or "content" not in data:
        return None
    try:
        return base64.b64decode(data["content"]).decode("utf-8", errors="ignore")
    except Exception:
        return None


# =============================================================================
# 3. HIGHER-LEVEL HELPERS FOR CHECKER SCRIPTS
# =============================================================================

# Filenames that count as a "root README" for prioritisation purposes.
# The set is lowercase so comparisons are case-insensitive.
ROOT_README_BASENAMES = {
    "readme.md",
    "readme.rst",
    "readme.txt",
    "readme",
}


def path_priority_with_readme_first(path: str, nested_readme_last: bool = False) -> int:
    """Return a sort key that prioritises root READMEs over deeper files.

    Used as ``key=`` for ``sorted()`` so that when we collect repository
    content for the LLM, the most informative files appear first and the
    less informative ones get truncated last when the character budget
    runs out.

    Priority levels (lower = appears first in the sorted list):

        0 → Root README files (readme.md, readme.rst, etc.).
        1 → Other root-level files (no "/" in the path).
        2 → Nested files (in any subdirectory).
        3 → Nested README files — only used when ``nested_readme_last=True``,
            which is helpful when nested READMEs are auto-generated boilerplate
            we want to consider last.

    Args:
        path: A file path relative to the repo root, e.g. ``"docs/intro.md"``.
        nested_readme_last: When True, push nested READMEs to priority 3.

    Returns:
        An integer 0–3 suitable for use as a sort key.
    """
    path_lower = path.lower()

    if path_lower in ROOT_README_BASENAMES:
        return 0
    if "/" not in path_lower:
        return 1
    if nested_readme_last and "readme" in path_lower:
        return 3
    return 2


def first_readme_path(all_files: Sequence[Mapping[str, Any]]) -> Optional[str]:
    """Return the path of the first root-level README in a GitHub file listing.

    Args:
        all_files: The list returned by ``list_all_repo_files``.

    Returns:
        str | None: The first path whose lowercased name starts with
        ``"readme"`` (e.g. ``"README.md"``, ``"readme.rst"``).  Returns
        None if no README is present.
    """
    for entry in all_files:
        path = str(entry.get("path", ""))
        if path.lower().startswith("readme"):
            return path
    return None


def fetch_paths_with_char_budget(
    owner: str,
    repo: str,
    file_paths: Sequence[str],
    max_total_chars: int,
    *,
    fetch_content: Callable[[str, str, str], Optional[str]],
    start_total_chars: int = 0,
    pause_seconds: float = 0.2,
    header_label: str = "FILE",
    per_file_char_cap: Optional[int] = None,
) -> Tuple[List[str], List[str], int]:
    """Fetch multiple repo files into one payload while respecting a char cap.

    LLM context windows are finite, so we need a way to assemble file content
    for the LLM that stops *before* we run out of space.  This helper:

      1. Iterates over ``file_paths`` in the order given (caller controls
         priority by sorting first).
      2. Calls ``fetch_content(owner, repo, path)`` on each.
      3. Truncates each file to ``per_file_char_cap`` (if set), so a single
         huge file cannot eat the entire budget.
      4. Stops as soon as ``max_total_chars`` is reached.
      5. Sleeps ``pause_seconds`` between requests to be polite to the
         GitHub API and stay under secondary rate limits.

    Each fetched file is wrapped with a header line so the LLM can tell
    files apart in the combined payload:

        ### FILE: src/preprocess.py
        <truncated file content>

    Args:
        owner, repo:        Identify the repo for ``fetch_content``.
        file_paths:         Paths to attempt, in priority order.
        max_total_chars:    Hard cap on combined output size.
        fetch_content:      Function ``(owner, repo, path) -> str | None``.
                            Usually ``fetch_file_content`` from this module,
                            but the parameter exists so callers can inject
                            a cached or mocked version.
        start_total_chars:  Initial budget already consumed (use this when
                            you call this helper multiple times with
                            different priority buckets).
        pause_seconds:      Sleep between API calls.
        header_label:       The label written before each file's content
                            (e.g. "PREPROCESS-NAMED" so the LLM knows
                            which bucket the file came from).
        per_file_char_cap:  Optional per-file truncation limit.

    Returns:
        A tuple ``(blocks, fetched_paths, total_chars)``:
          * ``blocks`` — list of formatted strings, one per file fetched.
            Caller typically does ``"\\n\\n".join(blocks)``.
          * ``fetched_paths`` — paths that were successfully fetched
            (subset of ``file_paths``).
          * ``total_chars`` — the running total, useful as
            ``start_total_chars`` for a subsequent call.
    """
    blocks: List[str] = []
    fetched: List[str] = []
    total_chars = max(0, start_total_chars)

    for file_path in file_paths:
        if total_chars >= max_total_chars:
            break

        content = fetch_content(owner, repo, file_path)
        if not content:
            continue

        remaining = max_total_chars - total_chars
        if remaining <= 0:
            break

        limit = remaining if per_file_char_cap is None else min(per_file_char_cap, remaining)
        snippet = content[:limit]

        blocks.append(f"### {header_label}: {file_path}\n{snippet}")
        fetched.append(file_path)
        total_chars += len(snippet)

        time.sleep(pause_seconds)

    return blocks, fetched, total_chars
