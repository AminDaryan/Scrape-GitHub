"""Shared utilities for confidence diagnosis and console reporting."""

from collections import Counter
from typing import Any, Callable, Dict, List, Mapping, Sequence


DEFAULT_UNKNOWN_DIAGNOSIS = {
    "id": "unknown",
    "label": "LLM was uncertain — no clear structural cause detected",
    "fix": "Manual review recommended",
}


def diagnose_with_rules(
    result: Mapping[str, Any],
    repo_content: str,
    diagnosis_rules: Sequence[Mapping[str, Any]],
    unknown_diagnosis: Mapping[str, Any] = DEFAULT_UNKNOWN_DIAGNOSIS,
) -> List[Mapping[str, Any]]:
    """Apply diagnosis rules safely and return all matching rule dictionaries."""
    diagnoses: List[Mapping[str, Any]] = []

    for rule in diagnosis_rules:
        check_fn = rule.get("check")
        if not callable(check_fn):
            continue
        try:
            if check_fn(result, repo_content):
                diagnoses.append(rule)
        except Exception:
            # Rule checks are heuristics; skip any check that raises.
            continue

    if not diagnoses:
        diagnoses.append(dict(unknown_diagnosis))

    return diagnoses


def print_confidence_report(
    results: Sequence[Mapping[str, Any]],
    repo_contents: Mapping[str, str],
    diagnose_result: Callable[[Mapping[str, Any], str], Sequence[Mapping[str, Any]]],
) -> None:
    """Print a confidence breakdown with root-cause diagnostics for non-high rows."""
    total = len(results)
    counts = Counter(r.get("confidence", "unknown") for r in results)
    high = counts.get("high", 0)
    medium = counts.get("medium", 0)
    low = counts.get("low", 0)

    denom = total if total else 1
    width = 90
    print("\n" + "═" * width)
    print("  CONFIDENCE REPORT")
    print("═" * width)
    print(f"  High   : {high:>3}  ({high/denom*100:.1f}%)")
    print(f"  Medium : {medium:>3}  ({medium/denom*100:.1f}%)")
    print(f"  Low    : {low:>3}  ({low/denom*100:.1f}%)")
    print(f"  Total  : {total}")

    non_high = [
        (row, repo_contents.get(row.get("title", ""), ""))
        for row in results
        if row.get("confidence") != "high"
    ]

    if not non_high:
        print("\n  All results are high confidence.")
        print("═" * width)
        return

    print(f"\n  {'─'*86}")
    print(f"  {'TITLE':<50} {'CONF':<8}  ROOT CAUSE")
    print(f"  {'─'*86}")

    for row, content in non_high:
        title_full = row.get("title", "")
        title = title_full[:48] + ("…" if len(title_full) > 48 else "")
        conf = row.get("confidence") or "?"
        diags = diagnose_result(row, content)

        for idx, diagnosis in enumerate(diags):
            if idx == 0:
                print(f"  {title:<50} {conf:<8}  CAUSE : {diagnosis['label']}")
                print(f"  {'':<50} {'':<8}  FIX   : {diagnosis['fix']}")
            else:
                print(f"  {'':<50} {'':<8}  ALSO  : {diagnosis['label']}")
        print(f"  {'─'*86}")

    print("═" * width)
