"""Shared helpers for parsing JSON-only LLM responses."""

import copy
import json
import re
from typing import Any, Dict, Iterable, Mapping, Optional


def parse_llm_json_response(
    raw_content: Optional[str],
    empty_payload: Mapping[str, Any],
    required_list_fields: Iterable[str] = (),
    legacy_single_to_list_fields: Optional[Mapping[str, str]] = None,
) -> Dict[str, Any]:
    """Parse an LLM JSON response with markdown-fence cleanup and fallbacks.

    - Returns a copy of empty_payload when response content is missing.
    - Strips optional ```json fences before decoding.
    - On parse failure, returns empty_payload with evidence updated.
    - Optionally maps legacy single-value keys into list keys.
    - Ensures required list fields exist in the final payload.
    """
    if not raw_content:
        return copy.deepcopy(dict(empty_payload))

    raw = raw_content.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        fallback = copy.deepcopy(dict(empty_payload))
        fallback["evidence"] = f"Parsing failed: {raw[:200]}"
        return fallback

    if legacy_single_to_list_fields:
        for old_key, new_key in legacy_single_to_list_fields.items():
            if old_key in parsed and new_key not in parsed:
                value = parsed.pop(old_key)
                parsed[new_key] = [value] if value else []

    for field in required_list_fields:
        if field not in parsed or parsed[field] is None:
            parsed[field] = []

    return parsed
