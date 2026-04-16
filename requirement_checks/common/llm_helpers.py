"""Shared LLM call, retry, and JSON response parsing for checker scripts."""

import copy
import json
import re
from typing import Any, Dict, Iterable, Mapping, Optional


# ---------------------------------------------------------------------------
# JSON response parsing
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# LLM call with retry
# ---------------------------------------------------------------------------


def llm_call_parse_retry(
    *,
    client,
    deployment,
    system_prompt,
    build_user_message,
    content,
    token_usage,
    empty_payload,
    required_list_fields=(),
    max_completion_tokens=1000,
    retry_truncate_chars=15_000,
    retry_max_tokens=None,
    preview_chars=500,
):
    """Send an LLM chat-completion, retry on empty, and parse the JSON result.

    *build_user_message(content_str)* builds the user prompt. If the first
    call returns nothing, content is truncated to *retry_truncate_chars* and
    retried once.  Returns the parsed JSON dict or a copy of *empty_payload*.
    """

    def _call(text, tokens):
        resp = client.chat.completions.create(
            model=deployment,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": build_user_message(text)},
            ],
            max_completion_tokens=tokens,
            response_format={"type": "json_object"},
        )
        token_usage.add_from_response(resp)
        return resp.choices[0].message.content

    raw = _call(content, max_completion_tokens)

    if not raw:
        truncated = content[:retry_truncate_chars] if content else ""
        print(
            f"\n  [empty response — retrying with {len(truncated):,} chars]",
            end=" ",
            flush=True,
        )
        raw = _call(truncated, retry_max_tokens or max_completion_tokens)

    if not raw:
        return dict(empty_payload)

    raw_preview = raw.strip()
    raw_preview = re.sub(r"^```(?:json)?\s*", "", raw_preview)
    raw_preview = re.sub(r"\s*```$", "", raw_preview)
    print("\nRAW LLM OUTPUT:\n", raw_preview[:preview_chars])

    return parse_llm_json_response(
        raw_content=raw,
        empty_payload=empty_payload,
        required_list_fields=required_list_fields,
    )
