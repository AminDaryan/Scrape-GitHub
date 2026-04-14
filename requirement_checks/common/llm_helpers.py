"""Shared LLM call-with-retry logic for checker scripts.

Extracts the common pattern: send chat-completion → retry on empty response
→ parse JSON via ``parse_llm_json_response``.
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
    """Send an LLM chat-completion, retry on empty, and parse the JSON result.

    *build_user_message(content_str)* builds the user prompt.  If the first
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
