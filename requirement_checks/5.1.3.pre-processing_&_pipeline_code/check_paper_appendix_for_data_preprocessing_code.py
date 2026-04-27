"""
Question 5.1.3 — Pre-processing & Pipeline Code classifier (repo-only)

Checks ONLY the GitHub repository for actual preprocessing / pipeline source
code (tokenization, filtering / deduplication / cleaning, or prompting-wrapper
construction). The paper text is no longer consulted.

Logic:
  1. Parse the GitHub URL from each paper entry. Non-GitHub repos are
     classified as NOT_APPLICABLE up front.                          → process()
  2. List every file in the repo via the GitHub API.                 → collect_repo_content()
  3. Sort files into three prioritised buckets:
       preprocessing-named source files (preprocess.py, tokenize.py …)
       files inside preprocessing-relevant folders (data/, pipeline/ …)
       other source files as a fallback.                             → collect_repo_content()
  4. Fetch the buckets within a shared character budget.             → collect_repo_content()
  5. Pass the combined text to the LLM and parse the verdict.        → process(), call_gpt()

Labels:
  EXISTS         — repo contains real preprocessing / pipeline code
  NOT_APPLICABLE — everything else (no code, only training/inference, or no GitHub URL)
"""

import json
import os
import re
import sys
import time
import urllib.parse
from collections import Counter
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

# Make both the parent package and the requirement_checks/ root importable,
# matching the layout used by check_github_repo_api_documentation.py.
_this = Path(__file__).resolve()
sys.path.insert(0, str(_this.parent.parent))
sys.path.insert(0, str(_this.parent.parent.parent))

from common.fetch_and_parse_github_repo import (
    parse_github_repo, list_all_repo_files, fetch_file_content,
)
from common.repo_content_helpers import fetch_paths_with_char_budget
from common.token_usage import TokenUsageTracker, print_token_usage_report

from openai_client import client, AZURE_OPENAI_DEPLOYMENT
from papers_from_database import PAPERS
from prompts import SYSTEM_PROMPT, USER_PROMPT


PAUSE        = 0.5
ALLOWED      = {"EXISTS", "NOT_APPLICABLE"}
TOKEN_USAGE  = TokenUsageTracker()
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")

# ---------------------------------------------------------------------------
# File-classification rules
# ---------------------------------------------------------------------------

# A file's basename matching this pattern is treated as preprocessing-relevant.
PREPROC_NAME_RE = re.compile(
    r"(preprocess|pre[_-]?proc|tokeniz|"
    r"clean(?:up|ing|er)?|filter|dedup(?:lic)?|"
    r"dataset|data[_-]?loader|dataloader|"
    r"prepare[_-]?data|prep[_-]?data|build[_-]?data|process[_-]?data|make[_-]?data|"
    r"prompt[_-]?(?:build|wrap|template)|pipeline|wrapper)",
    re.IGNORECASE,
)

# Any file under one of these top-level folders is treated as preprocessing-relevant.
PREPROC_FOLDER_PREFIXES = (
    "data/", "data_processing/", "data-processing/",
    "preprocessing/", "preprocess/",
    "pipeline/", "pipelines/",
    "tokenizer/", "tokenization/",
    "scripts/",
    "src/data/", "src/preprocessing/", "src/pipeline/",
)

SOURCE_EXTS = (
    ".py", ".sh", ".ipynb", ".js", ".ts",
    ".rb", ".go", ".jl", ".r", ".scala", ".rs",
)

# Folders we never look inside — vendored/test/build artefacts.
SKIP_FOLDER_PREFIXES = (
    "test/", "tests/", "__tests__/", ".github/",
    "node_modules/", "vendor/", "third_party/", "third-party/",
    "examples/", "docs/", "doc/",
    "build/", "dist/", ".cache/",
)

MAX_TOTAL_CHARS    = 80_000
PER_FILE_CHAR_CAP  = 8_000
MAX_NAMED_FILES    = 12
MAX_FOLDER_FILES   = 12
MAX_GENERIC_FILES  = 8

# ---------------------------------------------------------------------------
# Repo URL parsing
# ---------------------------------------------------------------------------

def parse_repo_url(url):
    """Extract (owner, repo) from a GitHub URL.

    Handles standard github.com URLs via the project's shared parser, and
    falls back to a regex for `<owner>.github.io/<repo>` GitHub Pages links
    that often appear as the project page in academic papers. Returns None
    for any URL that isn't on GitHub (HuggingFace, anonymous.4open.science,
    project websites, etc.) — those are skipped at the call site.
    """
    if not url:
        return None
    url = url.strip()

    try:
        result = parse_github_repo(url)
    except Exception:
        result = None

    if result:
        if isinstance(result, tuple) and len(result) >= 2:
            return (result[0], result[1])
        if isinstance(result, dict):
            owner = result.get("owner")
            repo  = result.get("repo")
            if owner and repo:
                return (owner, repo)

    # GitHub Pages URL: https://<owner>.github.io/<repo>
    m = re.match(r"https?://([^./]+)\.github\.io/([^/?#]+)", url)
    if m:
        return (m.group(1), m.group(2))

    return None

# ---------------------------------------------------------------------------
# Repo content collection
# ---------------------------------------------------------------------------

def _is_skippable(path_lower):
    return any(path_lower.startswith(p) for p in SKIP_FOLDER_PREFIXES)

def _is_source(path_lower):
    if "." not in path_lower:
        return False
    ext = "." + path_lower.rsplit(".", 1)[-1]
    return ext in SOURCE_EXTS

def collect_repo_content(owner, repo):
    """Fetch preprocessing-relevant files from a GitHub repo.

    Files are sorted into three priority buckets and fetched within a shared
    character budget so that named preprocessing files (preprocess.py,
    tokenize.py, …) are guaranteed to make it into the LLM's context before
    generic source code can crowd them out. Returns a (combined_text,
    fetched_paths) tuple; combined_text is empty if the repo is unreachable.
    """
    print(f"  [GitHub] Listing {owner}/{repo}")
    try:
        all_files = list_all_repo_files(owner, repo)
    except Exception as e:
        print(f"  [GitHub] ERROR during listing: {type(e).__name__}: {e}")
        return "", []

    if not all_files:
        print(f"  [GitHub] Repo empty or unreachable")
        return "", []

    print(f"  [GitHub] {len(all_files)} files in repo")

    named_paths   = []   # source files whose basename matches PREPROC_NAME_RE
    folder_paths  = []   # source files inside a preprocessing-relevant folder
    generic_paths = []   # other source files — fallback context

    for f in all_files:
        fpath     = f["path"] if isinstance(f, dict) else f
        fpath_low = fpath.lower()
        basename  = fpath_low.split("/")[-1]

        if _is_skippable(fpath_low):
            continue
        if not _is_source(fpath_low):
            continue

        if PREPROC_NAME_RE.search(basename):
            named_paths.append(fpath)
        elif any(fpath_low.startswith(p) for p in PREPROC_FOLDER_PREFIXES):
            folder_paths.append(fpath)
        else:
            generic_paths.append(fpath)

    # Sort by depth then name for determinism. Shallower paths come first
    # because top-level scripts are usually the canonical entry points.
    for lst in (named_paths, folder_paths, generic_paths):
        lst.sort(key=lambda p: (p.count("/"), p.lower()))

    named_paths   = named_paths[:MAX_NAMED_FILES]
    folder_paths  = folder_paths[:MAX_FOLDER_FILES]
    generic_paths = generic_paths[:MAX_GENERIC_FILES]

    print(f"  [GitHub] Selected: {len(named_paths)} named, "
          f"{len(folder_paths)} in folders, {len(generic_paths)} generic")

    combined    = []
    fetched     = []
    total_chars = 0

    for label, paths in (("PREPROCESS-NAMED",  named_paths),
                         ("PREPROCESS-FOLDER", folder_paths),
                         ("GENERIC SOURCE",    generic_paths)):
        if not paths:
            continue
        blocks, got, total_chars = fetch_paths_with_char_budget(
            owner, repo, paths, MAX_TOTAL_CHARS,
            fetch_content=fetch_file_content,
            start_total_chars=total_chars,
            pause_seconds=0.15,
            header_label=label,
            per_file_char_cap=PER_FILE_CHAR_CAP,
        )
        combined.extend(blocks)
        fetched.extend(got)

    return "\n\n".join(combined), fetched

# ---------------------------------------------------------------------------
# GPT call
# ---------------------------------------------------------------------------

def call_gpt(messages):
    """Send role-tagged messages to the configured Azure OpenAI deployment.

    Returns the raw response string. Empty string on error so the caller can
    fall through to the parser's default NOT_APPLICABLE without raising.
    """
    total_chars = sum(len(m["content"]) for m in messages)
    print(f"  [GPT] Calling {AZURE_OPENAI_DEPLOYMENT} — "
          f"{total_chars:,} total chars, max_completion_tokens=4000")
    try:
        resp = client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT,
            messages=messages,
            max_completion_tokens=4000,
        )
        TOKEN_USAGE.add_from_response(resp)
        raw = (resp.choices[0].message.content or "").strip()
        if not raw:
            print(f"  [GPT] ERROR: Empty response")
        else:
            print(f"  [GPT] OK — {len(raw)} chars: {raw[:200]!r}")
        return raw
    except Exception as e:
        print(f"  [GPT] ERROR: {type(e).__name__}: {e}")
        return ""

# ---------------------------------------------------------------------------
# Parse GPT response
# ---------------------------------------------------------------------------

def parse(raw):
    """Parse the LLM's JSON response and validate its classification field.

    Anything outside {EXISTS, NOT_APPLICABLE} is coerced to NOT_APPLICABLE so a
    malformed reply degrades safely instead of being silently mis-counted.
    """
    result = {
        "classification":       "NOT_APPLICABLE",
        "confidence":           0,
        "key_quotes":           [],
        "matched_categories":   [],
        "preprocessing_files":  [],
        "diagnostic_notes":     "N/A",
        "final_reason":         "N/A",
        "evidence":             "N/A",
    }
    if not raw:
        result["diagnostic_notes"] = "EMPTY GPT RESPONSE"
        return result

    try:
        data = json.loads(raw)
        result.update(data)
        result["evidence"] = (
            "; ".join(data.get("key_quotes", [])) or
            data.get("diagnostic_notes", "N/A")
        )
    except json.JSONDecodeError:
        result["diagnostic_notes"] = f"INVALID JSON — raw started: {raw[:100]}"

    if result.get("classification") not in ALLOWED:
        result["classification"] = "NOT_APPLICABLE"

    return result

# ---------------------------------------------------------------------------
# Per-paper pipeline
# ---------------------------------------------------------------------------

def process(entry):
    """Run the classification pipeline on a single paper entry's GitHub repo.

    Sources reported on the result dict:
      GITHUB     — repo was fetched and the LLM produced a verdict
      NO_GITHUB  — entry's repo URL is not a GitHub URL (HuggingFace, etc.)
      EMPTY_REPO — repo was reachable but yielded no source files
    """
    title    = entry.get("title", "")
    repo_url = entry.get("repo", "") or ""

    parsed = parse_repo_url(repo_url)
    if not parsed:
        print(f"  [PIPELINE] Not a GitHub URL: {repo_url!r}")
        return {
            "title":             title,
            "semanticscholarid": entry.get("semanticscholarid", ""),
            "repo":              repo_url,
            "classification":    "NOT_APPLICABLE",
            "evidence":          "Repository URL is not a GitHub URL",
            "reasoning":         "Skipped — non-GitHub repository",
            "source":            "NO_GITHUB",
            "files_checked":     [],
        }

    owner, repo = parsed
    print(f"  [PIPELINE] GitHub: {owner}/{repo}")
    content, fetched_paths = collect_repo_content(owner, repo)

    if not content:
        return {
            "title":             title,
            "semanticscholarid": entry.get("semanticscholarid", ""),
            "repo":              repo_url,
            "classification":    "NOT_APPLICABLE",
            "evidence":          "No source files retrievable from the repo",
            "reasoning":         "Empty or inaccessible repository",
            "source":            "EMPTY_REPO",
            "files_checked":     [],
        }

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user",   "content": USER_PROMPT.format(
            title=title, repo_url=repo_url, repo_content=content)},
    ]
    raw    = call_gpt(messages)
    result = parse(raw)

    result["title"]             = title
    result["semanticscholarid"] = entry.get("semanticscholarid", "")
    result["repo"]              = repo_url
    result["source"]            = "GITHUB"
    result["files_checked"]     = fetched_paths
    result["reasoning"]         = result.get("final_reason") or result.get("diagnostic_notes", "")
    return result

# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def save_excel(results, path):
    """Write per-paper predictions and a Summary sheet to an Excel workbook.

    The Results sheet colour-codes each row by predicted label and source.
    The Summary sheet shows label counts and a breakdown by source.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    COLOR = {
        "EXISTS":         "FFEB9C",
        "NOT_APPLICABLE": "FCE4D6",
    }
    LABEL = {
        "EXISTS":         "Exists",
        "NOT_APPLICABLE": "Not applicable",
    }
    SOURCE_COLOR = {
        "GITHUB":     "BDD7EE",
        "NO_GITHUB":  "F2F2F2",
        "EMPTY_REPO": "FFE4B5",
        "ERROR":      "FFC7CE",
    }

    wb = Workbook()
    ws = wb.active
    ws.title = "Results"

    headers = ["#", "Title", "Repo", "Prediction", "Source",
               "Files Checked", "Evidence", "Reasoning"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        c.fill      = PatternFill("solid", start_color="2F4F6F")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 25

    for i, r in enumerate(results, 1):
        row = i + 1
        clf = r.get("classification", "NOT_APPLICABLE")
        bg  = "F9F9F9" if i % 2 == 0 else "FFFFFF"

        def cell(col, val, bold=False, fg="000000", bg_ov=None):
            c = ws.cell(row=row, column=col, value=val)
            c.font      = Font(name="Arial", size=10, bold=bold, color=fg)
            c.alignment = Alignment(
                horizontal="center" if col in (1, 4, 5) else "left",
                vertical="top", wrap_text=True)
            c.fill = PatternFill("solid", start_color=bg_ov or bg)

        cell(1, i)
        cell(2, r.get("title", ""))
        cell(3, r.get("repo",  ""))

        c4 = ws.cell(row=row, column=4, value=LABEL.get(clf, clf))
        c4.font      = Font(bold=True, name="Arial", size=10)
        c4.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        c4.fill      = PatternFill("solid", start_color=COLOR.get(clf, "F2F2F2"))

        src = r.get("source", "")
        c5 = ws.cell(row=row, column=5, value=src)
        c5.font      = Font(name="Arial", size=10, bold=True)
        c5.alignment = Alignment(horizontal="center", vertical="top")
        c5.fill      = PatternFill("solid", start_color=SOURCE_COLOR.get(src, bg))

        cell(6, "\n".join(r.get("files_checked", [])))
        cell(7, r.get("evidence",  ""))
        cell(8, r.get("reasoning", ""))
        ws.row_dimensions[row].height = 60

    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Total papers"
    ws2["B1"] = len(results)
    ws2["A1"].font = ws2["B1"].font = Font(bold=True, name="Arial", size=12)

    for col, h in enumerate(["Label", "Count"], 1):
        ws2.cell(row=3, column=col, value=h).font = Font(bold=True, name="Arial")
    counts = Counter(r.get("classification", "NOT_APPLICABLE") for r in results)
    r2 = 4
    for lbl in ("EXISTS", "NOT_APPLICABLE"):
        ws2.cell(row=r2, column=1, value=LABEL.get(lbl, lbl))
        ws2.cell(row=r2, column=2, value=counts.get(lbl, 0))
        r2 += 1

    r2 += 1
    ws2.cell(row=r2, column=1, value="By source").font = Font(bold=True, name="Arial")
    r2 += 1
    sources = Counter(r.get("source", "") for r in results)
    for src, cnt in sources.most_common():
        ws2.cell(row=r2, column=1, value=src or "(none)")
        ws2.cell(row=r2, column=2, value=cnt)
        r2 += 1

    for col, w in [("A", 22), ("B", 14)]:
        ws2.column_dimensions[col].width = w
    for col, w in [("A", 5), ("B", 42), ("C", 38), ("D", 14),
                   ("E", 12), ("F", 35), ("G", 50), ("H", 50)]:
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(path)
    print(f"\n  Saved: {path}")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    results = []
    for entry in PAPERS:
        title = entry.get("title", "")
        print(f"\n{'='*60}")
        print(f"  {title[:70]}")
        print(f"{'='*60}")
        try:
            result = process(entry)
        except Exception as e:
            print(f"  [FATAL] {type(e).__name__}: {e}")
            result = {
                "title":             title,
                "semanticscholarid": entry.get("semanticscholarid", ""),
                "repo":              entry.get("repo", ""),
                "classification":    "NOT_APPLICABLE",
                "evidence":          f"ERROR: {e}",
                "reasoning":         "Exception during processing",
                "source":            "ERROR",
                "files_checked":     [],
            }
        results.append(result)
        print(f"  -> FINAL: [{result['source']}] {result['classification']}")
        time.sleep(PAUSE)

    print(f"\n{'='*60}\nSUMMARY\n{'='*60}")
    print(f"  Labels:  {dict(Counter(r['classification'] for r in results))}")
    print(f"  Sources: {dict(Counter(r['source']         for r in results))}")
    print_token_usage_report(TOKEN_USAGE, AZURE_OPENAI_DEPLOYMENT)
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results_513.xlsx")
    save_excel(results, out)