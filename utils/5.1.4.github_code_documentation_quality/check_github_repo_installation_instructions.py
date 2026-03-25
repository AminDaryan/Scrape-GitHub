import os, re, sys, time, base64, json
import urllib.parse, requests
from pathlib import Path
from collections import Counter
from urllib.request import urlopen, Request
from urllib.error import HTTPError, URLError
from dotenv import load_dotenv

load_dotenv()

sys.path.append(str(Path(__file__).resolve().parent.parent))
from openai_client import client, AZURE_OPENAI_DEPLOYMENT
from papers_from_database import PAPERS

GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")

# Files to fetch from each repo and send to the LLM
TARGET_FILENAMES = {
    "readme.md", "readme.rst", "readme.txt", "readme",
    "install.md", "install.rst", "install.txt",
    "setup.py", "setup.cfg", "pyproject.toml",
    "requirements.txt", "requirements-dev.txt",
    "environment.yml", "environment.yaml",
    "dockerfile", "docker-compose.yml", "docker-compose.yaml",
    "makefile",
    "getting_started.md", "getting-started.md",
    "docs/install.md", "docs/installation.md", "docs/setup.md",
}

# Max chars sent to LLM per repo (keeps token costs low)
MAX_CONTENT_CHARS = 10_000

# ─── GitHub API helpers ───────────────────────────────────────────────────────

def parse_github_repo(url):
    """Extract (owner, repo) from a GitHub URL, or (None, None)."""
    match = re.search(r"github\.com/([^/]+)/([^/?.#]+)", url)
    if match:
        owner = match.group(1)
        repo = match.group(2).rstrip("/")
        if repo.endswith(".git"):
            repo = repo[:-4]
        return owner, repo
    return None, None


def is_github(url):
    return "github.com" in url


def github_get(path):
    """Make a GitHub API GET request and return parsed JSON."""
    url = f"https://api.github.com/{path.lstrip('/')}"
    headers = {"Accept": "application/vnd.github.v3+json"}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    req = Request(url, headers=headers)
    try:
        with urlopen(req, timeout=10) as r:
            return json.loads(r.read().decode())
    except HTTPError as e:
        if e.code == 403:
            raise RuntimeError("GitHub API rate limit hit. Set GITHUB_TOKEN or wait an hour.")
        return None
    except URLError:
        return None


def list_all_repo_files(owner, repo):
    """
    Use the Git Trees API (recursive=1) to list every file in the repo in
    one request — far cheaper than walking the tree directory by directory.
    """
    data = github_get(f"repos/{owner}/{repo}/git/trees/HEAD?recursive=1")
    if not data or "tree" not in data:
        return []
    return [item for item in data["tree"] if item.get("type") == "blob"]


def fetch_file_content(owner, repo, file_path):
    """Fetch and base64-decode a single file's content via the Contents API."""
    data = github_get(f"repos/{owner}/{repo}/contents/{file_path}")
    if not data or "content" not in data:
        return None
    try:
        return base64.b64decode(data["content"]).decode("utf-8", errors="ignore")
    except Exception:
        return None


def collect_repo_content(owner, repo):
    """
    Fetch all targeted files from the repo and concatenate into a single
    string for the LLM. Returns (content_string, list_of_fetched_paths).
    """
    all_files = list_all_repo_files(owner, repo)
    if not all_files:
        return "", []

    # Select files whose path or basename matches our target set
    selected_paths = []
    for f in all_files:
        fpath_lower = f["path"].lower()
        basename = fpath_lower.split("/")[-1]
        if fpath_lower in TARGET_FILENAMES or basename in TARGET_FILENAMES:
            selected_paths.append(f["path"])

    # Fallback: grab just the root README if nothing matched
    if not selected_paths:
        for f in all_files:
            if f["path"].lower().startswith("readme"):
                selected_paths.append(f["path"])
                break

    fetched, combined, total_chars = [], [], 0
    for file_path in selected_paths:
        if total_chars >= MAX_CONTENT_CHARS:
            break
        content = fetch_file_content(owner, repo, file_path)
        if content:
            snippet = content[: MAX_CONTENT_CHARS - total_chars]
            combined.append(f"### FILE: {file_path}\n{snippet}")
            total_chars += len(snippet)
            fetched.append(file_path)
        time.sleep(0.2)  # be polite to the GitHub API

    return "\n\n".join(combined), fetched


# ─── LLM analysis ─────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are an expert code reviewer analysing GitHub repositories for academic papers.

Your task: determine whether the repository contains INSTALLATION INSTRUCTIONS —
i.e. any content that explains how to set up the environment, install dependencies,
or get the code running (pip install, conda, docker, apt-get, virtualenv, etc.).

You will receive the concatenated contents of key files (README, setup files,
requirements, Dockerfiles, Makefiles, etc.).

Respond with a JSON object ONLY — no extra text, no markdown fences:
{
  "has_installation_instructions": true | false,
  "confidence": "high" | "medium" | "low",
  "evidence": "<one sentence describing what you found or why you concluded no instructions exist>",
  "instruction_types": ["list", "of", "found", "types"]
}

Possible instruction_types values (use as many as apply):
  pip install, conda environment, docker setup, system packages,
  virtual environment, requirements file, makefile, manual build steps, other
If none found, use an empty list [].
Return ONLY valid JSON.
Do NOT include any explanation, text, or markdown.
Your entire response must be a single valid JSON object.
"""


def llm_check_installation(repo_content, paper_title):
    """Send repo content to the LLM and parse its structured response."""
    user_message = (
        f"Paper: {paper_title}\n\n"
        "Below are the contents of key files fetched from its GitHub repository.\n"
        "Does this repository contain installation instructions?\n\n"
        + (repo_content if repo_content else "[No relevant files found in the repository]")
    )

    response = client.chat.completions.create(
        model=AZURE_OPENAI_DEPLOYMENT,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_message},
        ],
        max_completion_tokens=1000,
        response_format={"type": "json_object"}   # 🔥 ADD THIS
    )

    content = response.choices[0].message.content

    if not content:
        return {
            "has_installation_instructions": None,
            "confidence": "low",
            "evidence": "Empty LLM response",
            "instruction_types": [],
        }

    raw = content.strip()
    # Strip markdown code fences if the model wraps the JSON
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    
    print("\nRAW LLM OUTPUT:\n", raw[:500])

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {
            "has_installation_instructions": None,
            "confidence": "low",
            "evidence": f"Parsing failed: {raw[:200]}",
            "instruction_types": [],
        }


# ─── Per-paper orchestration ──────────────────────────────────────────────────

def check_paper(paper):
    url = paper.get("repo", "")
    result = {
        "title":             paper["title"],
        "repo":              url,
        "status":            None,   # "yes" | "no" | "skipped" | "error"
        "confidence":        None,
        "evidence":          "",
        "instruction_types": [],
        "files_checked":     [],
        "note":              "",
    }

    if not is_github(url):
        result["status"] = "skipped"
        result["note"]   = "Not a GitHub repo"
        return result

    owner, repo = parse_github_repo(url)
    if not owner:
        result["status"] = "error"
        result["note"]   = "Could not parse GitHub URL"
        return result

    try:
        repo_content, files_checked = collect_repo_content(owner, repo)
        result["files_checked"] = files_checked

        llm_result = llm_check_installation(repo_content, paper["title"])

        result["status"]            = "yes" if llm_result.get("has_installation_instructions") else "no"
        result["confidence"]        = llm_result.get("confidence", "unknown")
        result["evidence"]          = llm_result.get("evidence", "")
        result["instruction_types"] = llm_result.get("instruction_types", [])
        
        if llm_result.get("has_installation_instructions") is None:
            result["status"] = "error"
        else:
            result["status"] = "yes" if llm_result["has_installation_instructions"] else "no"

    except RuntimeError as e:
        result["status"] = "error"
        result["note"]   = str(e)
    except Exception as e:
        result["status"] = "error"
        result["note"]   = f"Unexpected error: {e}"

    return result


# ─── Output helpers ───────────────────────────────────────────────────────────

def print_results(results):
    W = 110
    print("\n" + "=" * W)
    print(f"{'#':<4} {'STATUS':<10} {'CONF':<8} {'INSTRUCTION TYPES':<36} TITLE")
    print("=" * W)

    for i, r in enumerate(results, 1):
        icon = {
            "yes":     "YES",
            "no":      "NO",
            "skipped": "SKIP",
            "error":   "ERR",
        }.get(r["status"], "?")

        types_str = ", ".join(r.get("instruction_types", []))[:34]
        if not types_str and r.get("note"):
            types_str = r["note"][:34]

        title_short = r["title"][:52] + ("..." if len(r["title"]) > 52 else "")
        conf = r.get("confidence") or "-"
        print(f"{i:<4} {icon:<10} {conf:<8} {types_str:<36} {title_short}")

        if r.get("evidence"):
            print(f"       -> {r['evidence']}")
        if r.get("files_checked"):
            print(f"       files: {', '.join(r['files_checked'])}")

    print("=" * W)
    yes     = sum(1 for r in results if r["status"] == "yes")
    no      = sum(1 for r in results if r["status"] == "no")
    skipped = sum(1 for r in results if r["status"] == "skipped")
    errors  = sum(1 for r in results if r["status"] == "error")
    print(
        f"\nSUMMARY: {yes} have install instructions | "
        f"{no} missing | {skipped} skipped (non-GitHub) | {errors} errors\n"
    )



def save_results(results, path="installation_instructions_results.json"):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"Full results saved to {path}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print(f"Checking {len(PAPERS)} repos for installation instructions "
          f"using LLM ({AZURE_OPENAI_DEPLOYMENT})...")
    if not GITHUB_TOKEN:
        print("Tip: Set GITHUB_TOKEN in your .env to avoid the 60 req/hr GitHub rate limit.\n")

    results = []
    for i, paper in enumerate(PAPERS, 1):
        owner, repo_name = parse_github_repo(paper.get("repo", ""))
        label = f"{owner}/{repo_name}" if owner else paper["repo"]
        print(f"[{i:>2}/{len(PAPERS)}] {label} ...", end=" ", flush=True)

        result = check_paper(paper)
        results.append(result)

        icon = {"yes": "OK", "no": "NO", "skipped": "SKIP", "error": "ERR"}.get(result["status"], "?")
        extra = result.get("confidence") or result.get("note") or ""
        print(f"{icon}  {extra}")

        time.sleep(0.3)

    print_results(results)
    save_results(results)


if __name__ == "__main__":
    main()