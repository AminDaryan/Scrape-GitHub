"""Shared helpers for building repository content payloads for LLM checks."""

import time
from typing import Any, Callable, Mapping, Optional, Sequence, Tuple, List


ROOT_README_BASENAMES = {
    "readme.md",
    "readme.rst",
    "readme.txt",
    "readme",
}


def path_priority_with_readme_first(path: str, nested_readme_last: bool = False) -> int:
    """Return a sort key that prioritizes root README content first.

    Priority order:
    0) Root-level README files.
    1) Other root-level files.
    2) Nested files.
    3) Nested README files (only when nested_readme_last=True).
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
    """Return the first root README-like path from a GitHub file listing."""
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
    """Fetch multiple repository files while respecting a total char budget.

    Returns a tuple of:
    - content blocks formatted as "### <HEADER>: <path>\n<snippet>"
    - fetched paths
    - final total chars consumed
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
