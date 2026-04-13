"""Shared helpers for status aggregation and Excel-summary formatting.

This module centralizes tiny but repeated reporting logic that appears in
multiple checker scripts:
1) Mapping result statuses to fill colors for spreadsheet output.
2) Counting statuses in a robust, case-insensitive way.
3) Building a safe coverage formula string for Excel.

Keeping these in one place reduces duplicate code and ensures that reports from
different scripts stay consistent.
"""

from typing import Any, Dict, Mapping, Sequence


# Excel-compatible fill colors (RRGGBB, no leading '#').
# These keys also define the default "known" statuses used by count_statuses.
STATUS_FILL_COLORS = {
    "yes": "C6EFCE",      # positive finding
    "no": "FFDDC1",       # checked but missing
    "skipped": "FFEB9C",  # intentionally not evaluated
    "error": "FFC7CE",    # failed during processing
}


def count_statuses(
    results: Sequence[Mapping[str, Any]],
    status_key: str = "status",
) -> Dict[str, int]:
    """Count records by normalized status.

    Args:
        results: Sequence of result dictionaries.
        status_key: Dictionary key that stores each row's status value.

    Returns:
        A dictionary with counts for known statuses (yes/no/skipped/error) and
        any additional statuses encountered in the input.
    """
    # Pre-seed known statuses so downstream code can safely use .get("yes", 0)
    # or direct indexing without checking whether a key exists.
    counts: Dict[str, int] = {k: 0 for k in STATUS_FILL_COLORS}

    for row in results:
        # Normalize to avoid duplicate buckets such as "Yes", " YES ", etc.
        status = str(row.get(status_key, "")).strip().lower()
        if not status:
            continue
        # Preserve unknown statuses so anomalies stay visible in reports.
        counts[status] = counts.get(status, 0) + 1

    return counts


def coverage_formula(positive_count: int, total_count: int) -> str:
    # Build an Excel formula string for percentage coverage.
    denom = total_count if total_count else 1
    return f"={positive_count}/{denom}*100"
