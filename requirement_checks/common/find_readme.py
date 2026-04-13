"""README detection helper for GitHub repositories.

This module inspects a repository's full file tree through the GitHub Git Trees
API and returns a compact summary that can be written directly to reports.
"""

from typing import Mapping, Optional

import requests


README_PREFIX = "readme"
TREE_API_TEMPLATE = "https://api.github.com/repos/{owner}/{repo}/git/trees/HEAD?recursive=1"


def summarize_readme_files(
    owner: str,
    repo: str,
    headers: Optional[Mapping[str, str]] = None,
) -> str:
    """Return a README presence summary for a repository.

    Return contract:
    - "Unknown": GitHub API request failed or the payload was not usable.
    - "No": no README-like file was found.
    - "Yes (N): path1, path2, ...": one or more README-like files were found.

    A file is considered README-like if its basename starts with "readme"
    case-insensitively (for example README.md or readme.txt).
    """
    tree_url = TREE_API_TEMPLATE.format(owner=owner, repo=repo)
    request_headers = dict(headers) if headers else None

    try:
        response = requests.get(tree_url, headers=request_headers, timeout=15)
    except requests.RequestException:
        return "Unknown"

    if response.status_code != 200:
        return "Unknown"

    payload = response.json()
    tree = payload.get("tree", [])
    if not isinstance(tree, list):
        return "Unknown"

    readme_files = [
        item["path"]
        for item in tree
        if item.get("type") == "blob"
        and isinstance(item.get("path"), str)
        and item["path"].split("/")[-1].lower().startswith(README_PREFIX)
    ]

    if not readme_files:
        return "No"

    return f"Yes ({len(readme_files)}): " + ", ".join(readme_files)

