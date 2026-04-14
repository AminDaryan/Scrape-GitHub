"""Shared LLM call-with-retry logic for checker scripts.

Both the installation-instructions checker and the usage-examples checker
follow the same pattern:
  1. Build a user message from repository content.
  2. Send a chat-completion request.
  3. If the response is empty (context overflow), truncate and retry once.
  4. Parse the raw JSON output via parse_llm_json_response.

This module extracts that shared pattern into a single function.
"""

import re

from common.llm_response_parser import parse_llm_json_response


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
    """Send an LLM chat-completion request, retry on empty, and parse JSON.

    Args:
        client: OpenAI client instance.
        deployment: Model deployment name.
        system_prompt: System prompt string.
        build_user_message: Callable(content_str) -> user message string.
        content: Repository content string to analyse.
        token_usage: TokenUsageTracker instance for recording usage.
        empty_payload: Dict returned when both attempts produce no output.
        required_list_fields: Field names that must be lists in parsed result.
        max_completion_tokens: Token limit for the first attempt.
        retry_truncate_chars: Content character limit for the retry attempt.
        retry_max_tokens: Token limit for the retry (defaults to
            *max_completion_tokens*).
        preview_chars: How many characters of raw LLM output to print.

    Returns:
        Parsed JSON response dict.
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
