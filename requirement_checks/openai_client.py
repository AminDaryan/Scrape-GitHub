import json
import os
import litellm
from dotenv import load_dotenv

load_dotenv()

# Suppress litellm's verbose startup/request logging
litellm.suppress_debug_info = True
litellm.drop_params = True
os.environ.setdefault("LITELLM_LOG", "ERROR")

# ── Provider selection ─────────────────────────────────────────────────────────
# Set LLM_PROVIDER in .env to switch backends without touching code.
# Supported values: azure | openai | anthropic | mistral | ollama | gemini
LLM_PROVIDER = os.getenv("LLM_PROVIDER", "azure").lower()

# ── Azure OpenAI credentials (existing env var names, kept for compatibility) ──
OPENAI_KEY = os.getenv("OPENAI_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION")

# Map to the env var names litellm expects for Azure
if LLM_PROVIDER == "azure":
    if OPENAI_KEY:
        os.environ.setdefault("AZURE_API_KEY", OPENAI_KEY)
    if AZURE_OPENAI_ENDPOINT:
        os.environ.setdefault("AZURE_API_BASE", AZURE_OPENAI_ENDPOINT)
    if AZURE_OPENAI_API_VERSION:
        os.environ.setdefault("AZURE_API_VERSION", AZURE_OPENAI_API_VERSION)

# ── Deployment / model list ────────────────────────────────────────────────────
# For Azure: deployment names (e.g. "gpt-5-2025-08-07").
# For other providers: model names (e.g. "claude-3-5-sonnet-20241022").
# Supports JSON array or comma-separated: AZURE_OPENAI_DEPLOYMENT=["m1","m2"]
_raw_deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o")
try:
    _parsed = json.loads(_raw_deployment)
    if isinstance(_parsed, list):
        AZURE_OPENAI_DEPLOYMENTS = [str(d).strip() for d in _parsed if str(d).strip()]
    else:
        AZURE_OPENAI_DEPLOYMENTS = [str(_parsed).strip()]
except (json.JSONDecodeError, TypeError):
    AZURE_OPENAI_DEPLOYMENTS = [d.strip() for d in _raw_deployment.split(",") if d.strip()]

AZURE_OPENAI_DEPLOYMENT = AZURE_OPENAI_DEPLOYMENTS[0]  # backward compat


def _prefixed_model(name: str) -> str:
    """Return the litellm model string for the configured provider."""
    if "/" in name:
        return name  # already prefixed
    if LLM_PROVIDER == "azure":
        return f"azure/{name}"
    if LLM_PROVIDER in ("anthropic", "mistral", "ollama", "gemini"):
        return f"{LLM_PROVIDER}/{name}"
    return name  # openai and others: bare name is fine


# ── Thin wrapper that preserves client.chat.completions.create() interface ─────

class _Completions:
    def create(self, *, model, messages, max_completion_tokens=None,
               response_format=None, **kwargs):
        call_kwargs = dict(
            model=_prefixed_model(model),
            messages=messages,
            **kwargs,
        )
        if max_completion_tokens is not None:
            call_kwargs["max_completion_tokens"] = max_completion_tokens
        if response_format is not None:
            call_kwargs["response_format"] = response_format
        return litellm.completion(**call_kwargs)


class _Chat:
    def __init__(self):
        self.completions = _Completions()


class _LiteLLMClient:
    def __init__(self):
        self.chat = _Chat()


client = _LiteLLMClient()
