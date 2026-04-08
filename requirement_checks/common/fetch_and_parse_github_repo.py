
import os
import re
import sys
import base64
import json
import urllib.parse
from pathlib import Path
from urllib.request import urlopen, Request
from urllib.error import HTTPError, URLError
from dotenv import load_dotenv


load_dotenv()


# =============================================================================
# IMPORTS — papers list and LLM client
# =============================================================================
sys.path.append(str(Path(__file__).resolve().parent.parent))
from papers_from_database import PAPERS


GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")

try:
    sys.path.append(str(Path(__file__).resolve().parent.parent.parent))
    from openai_client import client, AZURE_OPENAI_DEPLOYMENT as DEPLOYMENT
except ImportError:
    import openai
    client     = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY", ""))
    DEPLOYMENT = os.getenv("OPENAI_MODEL", "gpt-4o")


# =============================================================================
# GITHUB API HELPERS
# =============================================================================
    
def parse_github_repo(url):
    """
    Extract (owner, repo) from a GitHub URL.
    Returns (None, None) if the URL is not a recognisable GitHub repo URL.
    """
    match = re.search(r"github\.com/([^/]+)/([^/?.#]+)", url)
    if match:
        owner = match.group(1)
        repo  = match.group(2).rstrip("/")
        if repo.endswith(".git"):
            repo = repo[:-4]
        return owner, repo
    return None, None


def is_github(url):
    return "github.com" in url


def github_get(path):
    """
    GET a GitHub REST API endpoint and return parsed JSON, or None on error.
    Raises RuntimeError when the rate limit is hit so the caller can bail out.
    """
    url     = f"https://api.github.com/{path.lstrip('/')}"
    headers = {"Accept": "application/vnd.github.v3+json"}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    req = Request(url, headers=headers)
    try:
        with urlopen(req, timeout=10) as r:
            return json.loads(r.read().decode())
    except HTTPError as e:
        if e.code == 403:
            raise RuntimeError(
                "GitHub API rate limit hit. Set GITHUB_TOKEN in .env or wait an hour."
            )
        return None
    except URLError:
        return None
    

def list_all_repo_files(owner, repo):
    """
    Return every file blob in the repo using the recursive Git Trees API.
    This is a single API call regardless of repo size.
    """
    data = github_get(f"repos/{owner}/{repo}/git/trees/HEAD?recursive=1")
    if not data or "tree" not in data:
        return []
    return [item for item in data["tree"] if item.get("type") == "blob"]


def fetch_file_content(owner, repo, file_path):
    """Download one file and return its UTF-8 text, or None on failure."""
    encoded_path = urllib.parse.quote(file_path, safe="/")  # ← encode spaces etc.
    data = github_get(f"repos/{owner}/{repo}/contents/{encoded_path}")
    if not data or "content" not in data:
        return None
    try:
        return base64.b64decode(data["content"]).decode("utf-8", errors="ignore")
    except Exception:
        return None