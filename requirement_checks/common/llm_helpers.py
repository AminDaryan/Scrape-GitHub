"""LLM helpers — call the model, track tokens, and parse the JSON response.

Every checker script in this project follows the same loop:

    1. Build a prompt that includes the paper's repo content.
    2. Send it to an LLM and get back a JSON object.
    3. Track how many tokens were spent (so we can report cost).
    4. Parse the JSON, with sensible fallbacks when things go wrong
       (empty response, malformed JSON, missing fields, …).

This module bundles the three helpers that handle steps 2–4:

  * :func:`llm_call_parse_retry` — the one-shot "send + retry + parse"
    function.  This is what checker scripts actually call.

  * :func:`parse_llm_json_response` — the lower-level JSON parser used
    inside ``llm_call_parse_retry``.  Exposed because some callers need
    to parse a response they obtained themselves.

  * :class:`TokenUsageTracker` — accumulates prompt/completion/total
    tokens across every LLM call in a script run.  Pass one instance
    into ``llm_call_parse_retry`` and it will be updated in place.

  * :func:`print_token_usage_report` — pretty-print the tracker totals
    at the end of a run.

Why JSON?
---------
All prompts in this project ask the model to return a single JSON object
(a "structured output").  This makes parsing reliable and keeps prompt
templates short.  When using OpenAI/Azure, we also pass
``response_format={"type": "json_object"}`` to force the API to enforce
that constraint.
"""

import copy
import json
import re
from dataclasses import dataclass
from typing import Any, Dict, Iterable, Mapping, Optional, Tuple


# =============================================================================
# TOKEN USAGE TRACKING
# =============================================================================

@dataclass
class TokenUsageTracker:
    """Accumulates LLM token usage across multiple calls in one run.

    Pass a single instance into every ``llm_call_parse_retry`` call you
    make for a given model, then print it at the end with
    :func:`print_token_usage_report`.  This lets you see — per model —
    how many requests you sent and how much input/output you spent.

    Attributes:
        request_count:        How many LLM API calls were made.
        prompt_tokens:        Total input tokens across all calls.
        completion_tokens:    Total output tokens across all calls.
        total_tokens:         Sum of input + output (sometimes the API
                              returns this directly; otherwise we compute
                              it from prompt + completion).
        missing_usage_count:  Number of responses that did NOT contain
                              token-usage metadata.  This usually means
                              the provider doesn't report it (e.g. some
                              local models) — the call still happened
                              but we couldn't measure it.
    """

    request_count: int = 0
    prompt_tokens: int = 0
    completion_tokens: int = 0
    total_tokens: int = 0
    missing_usage_count: int = 0

    def add_from_response(self, response: Any) -> Tuple[int, int, int]:
        """Extract usage from an LLM response object and add it to the totals.

        Different SDKs expose usage differently:
          * OpenAI Python SDK → ``response.usage`` (a Pydantic model).
          * Azure OpenAI / litellm → same as OpenAI in practice.
          * Some libraries return a plain dict with a ``"usage"`` key.

        This method tries each shape in turn so callers don't have to
        worry about the underlying SDK.  Returns the tuple
        ``(prompt, completion, total)`` for the *just-added* call so
        the caller can also log per-call usage if it wants to.
        """
        self.request_count += 1

        # 1. Attribute access (most common: SDK objects).
        usage = getattr(response, "usage", None)

        # 2. Dict-style access (when the SDK returns plain JSON).
        if usage is None and isinstance(response, Mapping):
            usage = response.get("usage")

        # 3. Pydantic-style: try .model_dump() and look in the result.
        if usage is None and hasattr(response, "model_dump"):
            try:
                dumped = response.model_dump()
                if isinstance(dumped, Mapping):
                    usage = dumped.get("usage")
            except Exception:
                usage = None

        # No luck — count it and move on.  Some providers genuinely don't
        # report usage; we don't want to crash a run for that.
        if usage is None:
            self.missing_usage_count += 1
            return 0, 0, 0

        # GPT-style providers use prompt_tokens/completion_tokens.
        # Anthropic uses input_tokens/output_tokens.  Try both.
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
    """Print a compact human-readable token-usage report to stdout.

    Designed to be called at the end of a checker run.  Keeps the output
    short so it can sit between the per-model summary and the next model's
    output without overwhelming the console.
    """

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
    """Try each key in order and return the first integer-like value found.

    Token usage fields are sometimes attributes (SDK objects), sometimes
    dict keys, and sometimes strings.  This helper handles all three so
    ``add_from_response`` doesn't need three separate code paths.

    Returns 0 if no key matches or the value is not an integer-like type.
    Booleans are explicitly skipped because ``isinstance(True, int)`` is
    True in Python and we don't want a stray boolean to be read as 1.
    """
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


# =============================================================================
# JSON RESPONSE PARSING
# =============================================================================

def parse_llm_json_response(
    raw_content: Optional[str],
    empty_payload: Mapping[str, Any],
    required_list_fields: Iterable[str] = (),
    legacy_single_to_list_fields: Optional[Mapping[str, str]] = None,
) -> Dict[str, Any]:
    """Parse an LLM JSON response with markdown-fence cleanup and fallbacks.

    LLMs sometimes wrap their JSON output in ```json ... ``` fences even
    when told not to.  Other times they return malformed JSON, or omit
    fields that the parser expects to be lists.  This function handles
    all three cases so callers always get back a *valid* dict matching
    the expected schema.

    Args:
        raw_content:
            The raw string returned by the LLM.  May be None if the API
            call failed or returned an empty response.
        empty_payload:
            A "default" dict with the expected schema and sensible empty
            values.  Returned (deep-copied) when ``raw_content`` is empty
            or the JSON cannot be parsed.  Must contain every key the
            caller expects to read.
        required_list_fields:
            Field names that must exist as a list in the final result.
            Any missing or null-valued field gets an empty list.  This
            avoids ``None`` surprises in callers that iterate over the
            field with ``for x in result["files"]``.
        legacy_single_to_list_fields:
            Optional mapping of ``{old_singular_key: new_list_key}`` for
            migrating legacy schema where a single string was returned
            instead of a list of strings.  When the old key is present
            and the new key is not, the value is wrapped in a list.

    Returns:
        A dict matching ``empty_payload``'s schema.  Never returns None.
        On parse failure, the dict will have ``evidence`` set to a snippet
        of the unparseable response so debugging is possible.
    """
    if not raw_content:
        return copy.deepcopy(dict(empty_payload))

    # Strip ```json ... ``` fences if the LLM ignored our "no fences"
    # instruction.  The leading fence may also include language.
    raw = raw_content.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        # Malformed JSON.  Return the empty payload but preserve enough
        # of the bad output that debugging is possible from the result.
        fallback = copy.deepcopy(dict(empty_payload))
        fallback["evidence"] = f"Parsing failed: {raw[:200]}"
        return fallback

    # Migrate any legacy single-value keys to their new list-value forms.
    if legacy_single_to_list_fields:
        for old_key, new_key in legacy_single_to_list_fields.items():
            if old_key in parsed and new_key not in parsed:
                value = parsed.pop(old_key)
                parsed[new_key] = [value] if value else []

    # Force every "this should be a list" field to actually be a list.
    for field in required_list_fields:
        if field not in parsed or parsed[field] is None:
            parsed[field] = []

    return parsed


# =============================================================================
# LLM CALL WITH RETRY + PARSE
# =============================================================================

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
    """Send a chat-completion request, retry on empty response, parse JSON.

    This is the top-level helper used by every checker script.  It wraps:
      * Building the message list (system + user).
      * Forcing ``response_format=json_object`` so the model returns valid JSON.
      * Recording token usage in ``token_usage``.
      * One automatic retry with truncated content if the first response
        is empty (a common failure mode when the input is too long for
        the model's context window).
      * Parsing the JSON via :func:`parse_llm_json_response`.

    Args:
        client:
            LLM client object exposing ``client.chat.completions.create(...)``.
            Compatible with the OpenAI/Azure SDKs and our litellm wrapper.
        deployment:
            Model name / Azure deployment name to send the request to.
        system_prompt:
            The system message — usually checker-specific policy.
        build_user_message:
            A callable ``(content_str) -> str`` that wraps the content
            in the checker-specific user message template.  Receiving a
            callable (rather than the final string) lets us call it twice:
            once with the full content, once with truncated content if
            we need to retry.
        content:
            The repo content (joined file blocks) to embed in the prompt.
        token_usage:
            A :class:`TokenUsageTracker` that gets updated with each call's
            prompt/completion/total counts.
        empty_payload:
            Fallback dict returned when the LLM returns an empty response
            even after retry.  See :func:`parse_llm_json_response`.
        required_list_fields:
            Fields that must be present as lists in the parsed result.
        max_completion_tokens:
            Max output tokens for the first attempt.  Must be large enough
            to fit the JSON the LLM is asked to produce.
        retry_truncate_chars:
            On retry, content is sliced to this many characters.  An empty
            response usually means the input exceeded the model's context
            window, so we make the input dramatically shorter and try once
            more.
        retry_max_tokens:
            Output-token budget for the retry.  Defaults to
            ``max_completion_tokens`` when None.
        preview_chars:
            Number of characters of the raw response to print to the
            console for debugging.

    Returns:
        dict: The parsed JSON response, or a copy of ``empty_payload``
        if both attempts returned empty.
    """

    def _call(text, tokens):
        """Inner: actually send the request and record token usage."""
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

    # First attempt with full content.
    raw = _call(content, max_completion_tokens)

    # Empty response usually means context-window overflow.  Retry once
    # with a shorter content slice.
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

    # Print a preview of the raw output so a developer running the script
    # can sanity-check the model's behaviour.
    raw_preview = raw.strip()
    raw_preview = re.sub(r"^```(?:json)?\s*", "", raw_preview)
    raw_preview = re.sub(r"\s*```$", "", raw_preview)
    print("\nRAW LLM OUTPUT:\n", raw_preview[:preview_chars])

    return parse_llm_json_response(
        raw_content=raw,
        empty_payload=empty_payload,
        required_list_fields=required_list_fields,
    )
