import os
from openai import AzureOpenAI
from dotenv import load_dotenv

# ── Load configuration from .env ───────────────────────────────────────────────
load_dotenv()  # take environment variables from .env


# Fetch values from environment
OPENAI_KEY = os.getenv("OPENAI_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION")
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o")  # default deployment


# Create the Azure OpenAI client
client = AzureOpenAI(
  api_key=OPENAI_KEY,
  azure_endpoint=AZURE_OPENAI_ENDPOINT,
  api_version= AZURE_OPENAI_API_VERSION
)

