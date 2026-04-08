"""Shared utilities for tracking LLM token usage."""

from dataclasses import dataclass
from typing import Any, Mapping, Tuple


@dataclass
class TokenUsageTracker:
    """Accumulates token usage across multiple LLM calls."""

    request_count: int = 0
    prompt_tokens: int = 0
    completion_tokens: int = 0
    total_tokens: int = 0
    missing_usage_count: int = 0

    def add_from_response(self, response: Any) -> Tuple[int, int, int]:
        """Extract usage from an SDK response object and add it to totals."""
        self.request_count += 1

        usage = getattr(response, "usage", None)
        if usage is None and isinstance(response, Mapping):
            usage = response.get("usage")

        if usage is None and hasattr(response, "model_dump"):
            try:
                dumped = response.model_dump()
                if isinstance(dumped, Mapping):
                    usage = dumped.get("usage")
            except Exception:
                usage = None

        if usage is None:
            self.missing_usage_count += 1
            return 0, 0, 0

        prompt = _extract_int(usage, ("prompt_tokens", "input_tokens"))
        completion = _extract_int(usage, ("completion_tokens", "output_tokens"))
        total = _extract_int(usage, ("total_tokens",))
        if total <= 0:
            total = prompt + completion

        self.prompt_tokens += prompt
        self.completion_tokens += completion
        self.total_tokens += total
        return prompt, completion, total


def print_token_usage_report(tracker: TokenUsageTracker, model_name: str = "") -> None:
    """Print a compact token-only report for a script run."""

    print("\n" + "=" * 78)
    print("LLM TOKEN USAGE")
    print("=" * 78)
    if model_name:
        print(f"Model/Deployment  : {model_name}")
    print(f"Requests           : {tracker.request_count}")
    print(f"Input tokens       : {tracker.prompt_tokens:,}")
    print(f"Output tokens      : {tracker.completion_tokens:,}")
    print(f"Total tokens       : {tracker.total_tokens:,}")

    if tracker.missing_usage_count:
        print(
            "Missing usage meta : "
            f"{tracker.missing_usage_count} response(s) did not include usage fields"
        )

    print("=" * 78)


def _extract_int(source: Any, keys: Tuple[str, ...]) -> int:
    for key in keys:
        value = None
        if isinstance(source, Mapping):
            value = source.get(key)
        else:
            value = getattr(source, key, None)

        if isinstance(value, bool):
            continue
        if isinstance(value, (int, float)):
            return int(value)
        if isinstance(value, str):
            text = value.strip()
            if text.isdigit():
                return int(text)
    return 0
