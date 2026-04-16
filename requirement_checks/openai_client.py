import json
import os
from openai import AzureOpenAI
from dotenv import load_dotenv

# ── Load configuration from .env ───────────────────────────────────────────────
load_dotenv()  # take environment variables from .env


# Fetch values from environment
OPENAI_KEY = os.getenv("OPENAI_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION")

# Support JSON array (e.g. '["gpt-5","gpt-5-mini"]') or comma-separated list
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


# Create the Azure OpenAI client
client = AzureOpenAI(
  api_key=OPENAI_KEY,
  azure_endpoint=AZURE_OPENAI_ENDPOINT,
  api_version= AZURE_OPENAI_API_VERSION
)

