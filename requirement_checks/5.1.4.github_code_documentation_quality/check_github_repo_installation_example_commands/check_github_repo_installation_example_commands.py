# =============================================================================
# usage_examples_checker.py
#
# PURPOSE:
#   For each academic paper in your database, this script visits the paper's
#   GitHub repository and finds ALL usage examples present (Jupyter notebooks,
#   demo scripts, README code snippets, tutorial files, etc.).
#
#   Each found example is recorded individually with:
#     - its file path inside the repo
#     - the example type (notebook / example script / README code snippet / …)
#     - a one-line description of what it demonstrates
#     - a direct clickable GitHub link
#
# HOW IT WORKS — high-level flow:
#   1. Read the list of papers (title + GitHub URL) from papers_from_database.py
#   2. For each repo, download the README, docs, ALL notebooks, and any
#      example/tutorial/demo files via the GitHub API (no cloning needed).
#   3. Send that content to an LLM and ask it to list EVERY usage example found.
#   4. If the LLM is unsure, fetch more files and retry automatically.
#   5. Save two Excel sheets:
#        "Results"      — one row per paper (summary + all example links)
#        "All Examples" — one row per individual example file (for easy browsing)
#
# REQUIREMENTS:
#   pip install openai openpyxl python-dotenv
#
# ENVIRONMENT VARIABLES (put these in a .env file next to this script):
#   GITHUB_TOKEN   — optional but strongly recommended (raises GitHub rate limit
#                    from 60 to 5,000 API requests per hour).
#   OPENAI_API_KEY — only needed if you are NOT using Azure OpenAI.
#   OPENAI_MODEL   — only needed if you are NOT using Azure OpenAI.
# =============================================================================

import os
import re
import sys
import time
import json
from pathlib import Path
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
# fetch from file fetch_&_parse_github_repo.py, which contains the GitHub API helpers and LLM call wrapper
from utils.common.fetch_and_parse_github_repo import (
    load_dotenv, parse_github_repo, is_github, list_all_repo_files, fetch_file_content
)
from utils.common.llm_response_parser import parse_llm_json_response
from utils.common.confidence_reporting import (
    diagnose_with_rules,
    print_confidence_report as print_shared_confidence_report,
)
from utils.common.token_usage import TokenUsageTracker, print_token_usage_report
from utils.common.result_status import (
    STATUS_FILL_COLORS,
    count_statuses,
    coverage_formula,
)
from utils.common.repo_content_helpers import (
    fetch_paths_with_char_budget,
    path_priority_with_readme_first,
)

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
# CONFIGURATION
# =============================================================================

# Named documentation files we always want to fetch (compared in lowercase).
TARGET_FILENAMES = {
    "readme.md", "readme.rst", "readme.txt", "readme",
    "usage.md", "usage.rst", "usage.txt",
    "quickstart.md", "quick_start.md", "quick-start.md",
    "tutorial.md", "tutorials.md",
    "examples.md", "example.md",
    "demo.md",
    "getting_started.md", "getting-started.md",
    "docs/usage.md", "docs/quickstart.md", "docs/tutorial.md",
    "docs/examples.md", "docs/demo.md", "docs/getting_started.md",
    "pyproject.toml", "setup.cfg",
}

# File extensions that are always usage examples by definition — we fetch ALL
# of them regardless of where they live in the repo (not capped like before).
# NOTE: .rmd and .qmd are excluded — they are analysis documents, not tutorials.
USAGE_EXTENSIONS = {".ipynb"}

# Extensions of plain example/demo scripts we also want to collect
SCRIPT_EXTENSIONS = {".py", ".r", ".sh", ".bash", ".m", ".jl"}

# Folder names whose entire contents are example/usage material
EXAMPLE_FOLDER_PREFIXES = ("examples/", "tutorials/", "demo/", "demos/",
                            "notebooks/", "notebook/", "scripts/", "sample/", "samples/")

# Maximum characters sent to the LLM per request
MAX_CONTENT_CHARS = 400_000

# Maximum number of notebooks to fully fetch (summarised) in one pass.
# We raised this from 5 → 20 so fewer notebooks are "skipped".
MAX_NOTEBOOKS = 20

# Maximum number of example scripts to fetch per pass
MAX_SCRIPTS = 15
TOKEN_USAGE = TokenUsageTracker()


# =============================================================================
# GITHUB API HELPERS
# =============================================================================


def summarise_notebook(raw_json_text, max_chars=8_000):
    """
    Distil a Jupyter notebook JSON into a compact plain-text block containing
    only code cells, markdown cells, and short text outputs.

    This strips base64-encoded images, widget state, and kernel metadata that
    would waste context-window tokens without helping the LLM.
    """
    try:
        nb = json.loads(raw_json_text)
    except Exception:
        return raw_json_text[:max_chars]

    lines = []
    for cell in nb.get("cells", []):
        ct  = cell.get("cell_type", "")
        src = "".join(cell.get("source", []))
        if not src.strip():
            continue
        if ct == "code":
            lines.append(f"[CODE]\n{src}")
            for out in cell.get("outputs", []):
                text = "".join(out.get("text", []))
                if text and len(text) < 500:
                    lines.append(f"[OUTPUT]\n{text}")
        elif ct == "markdown":
            lines.append(f"[MARKDOWN]\n{src}")

    return "\n\n".join(lines)[:max_chars]


def collect_repo_content(owner, repo):
    """
    Download all files that could contain usage examples and combine them into
    one large text string for the LLM.

    What we collect (in priority order so the most important files get the most
    character budget):
        1. Root README
        2. Other named documentation files (usage.md, quickstart.md, …)
        3. ALL markdown/rst files inside docs/, examples/, tutorials/, demo/, …
        4. ALL notebooks (.ipynb / .Rmd / .qmd) — up to MAX_NOTEBOOKS
        5. Example/demo Python/shell scripts — up to MAX_SCRIPTS

    Returns:
        content_string   — everything joined into one big string for the LLM
        fetched_paths    — list of paths successfully downloaded
        all_example_meta — list of dicts describing every example-like file
                           found in the repo tree (even ones not fetched due
                           to the char budget), used later to build links.
                           Each dict: {"path": str, "kind": "notebook"|"script"|"doc"}
    """
    all_files = list_all_repo_files(owner, repo)
    if not all_files:
        return "", [], []

    docs_paths     = []   # named documentation files
    notebook_paths = []   # .ipynb / .Rmd / .qmd files
    script_paths   = []   # example/demo scripts

    for f in all_files:
        fpath     = f["path"]
        fpath_low = fpath.lower()
        basename  = fpath_low.split("/")[-1]
        ext       = "." + basename.rsplit(".", 1)[-1] if "." in basename else ""

        # Notebooks anywhere in the tree
        if ext in USAGE_EXTENSIONS:
            notebook_paths.append(fpath)

        # Named doc files
        elif fpath_low in TARGET_FILENAMES or basename in TARGET_FILENAMES:
            docs_paths.append(fpath)

        # Markdown inside special folders
        elif any(fpath_low.startswith(p) for p in EXAMPLE_FOLDER_PREFIXES):
            if ext in (".md", ".rst", ".txt"):
                docs_paths.append(fpath)
            elif ext in SCRIPT_EXTENSIONS:
                script_paths.append(fpath)

        # Script files whose name suggests they are examples
        elif ext in SCRIPT_EXTENSIONS and any(
            kw in basename for kw in
            ["example", "demo", "tutorial", "quickstart", "sample", "run_", "test_usage"]
        ):
            script_paths.append(fpath)

    # Sort docs: root README first, then other root-level files, then nested
    docs_paths.sort(key=path_priority_with_readme_first)

    # Sort notebooks and scripts: shallowest first (root-level ones are most
    # likely to be the "main" usage example)
    notebook_paths.sort(key=lambda p: (p.count("/"), p.lower()))
    script_paths.sort(key=lambda p: (p.count("/"), p.lower()))

    # Build the metadata list for ALL example-like files (even unfetched ones),
    # so we can produce GitHub links for every one of them later.
    all_example_meta = (
        [{"path": p, "kind": "notebook"} for p in notebook_paths] +
        [{"path": p, "kind": "script"}   for p in script_paths]
    )

    # Cap how many we actually fetch to stay within the character budget
    notebooks_to_fetch = notebook_paths[:MAX_NOTEBOOKS]
    scripts_to_fetch   = script_paths[:MAX_SCRIPTS]

    fetched     = []
    combined    = []
    total_chars = 0

    # ── 1. Fetch documentation files ─────────────────────────────────────────
    doc_blocks, fetched_docs, total_chars = fetch_paths_with_char_budget(
        owner,
        repo,
        docs_paths,
        MAX_CONTENT_CHARS,
        fetch_content=fetch_file_content,
        start_total_chars=total_chars,
        pause_seconds=0.15,
        header_label="FILE",
    )
    combined.extend(doc_blocks)
    fetched.extend(fetched_docs)

    # ── 2. Fetch and summarise notebooks ─────────────────────────────────────
    remaining     = MAX_CONTENT_CHARS - total_chars
    budget_per_nb = min(8_000, remaining // max(len(notebooks_to_fetch), 1))

    for nb_path in notebooks_to_fetch:
        if total_chars >= MAX_CONTENT_CHARS:
            break
        raw = fetch_file_content(owner, repo, nb_path)
        if raw:
            summary = summarise_notebook(raw, max_chars=budget_per_nb)
            combined.append(f"### NOTEBOOK: {nb_path}\n{summary}")
            total_chars += len(summary)
            fetched.append(nb_path)
        time.sleep(0.15)

    # ── 3. Fetch example scripts ──────────────────────────────────────────────
    script_blocks, fetched_scripts, total_chars = fetch_paths_with_char_budget(
        owner,
        repo,
        scripts_to_fetch,
        MAX_CONTENT_CHARS,
        fetch_content=fetch_file_content,
        start_total_chars=total_chars,
        pause_seconds=0.15,
        header_label="SCRIPT",
        per_file_char_cap=4_000,
    )
    combined.extend(script_blocks)
    fetched.extend(fetched_scripts)

    return "\n\n".join(combined), fetched, all_example_meta


# =============================================================================
# LLM PROMPTS
# =============================================================================
# KEY CHANGE: every prompt now asks for `example_files` — a LIST of objects,
# one per example found — instead of a single `example_file` string.
# Each object has: path, type, description.

SYSTEM_PROMPT = """\
You are an expert code reviewer analysing GitHub repositories for academic papers.

Your task: find ALL usage examples in the repository — content that shows HOW
TO USE the code or model, beyond just installing it.

Usage examples include:
  - Jupyter notebooks (.ipynb)
  - Example or demo scripts (example.py, demo.py, run_example.sh, …)
  - Code snippets in the README that demonstrate the API or CLI
  - A "Quick-start" or "Usage" section in the README containing runnable code
  - Command-line usage examples with flags/arguments shown
  - Tutorial files or a tutorials/ / examples/ / demo/ folder

Does NOT count as a usage example:
  - Installation instructions alone (pip install X, conda env create, conda activate)
  - Navigation commands alone (cd src, mkdir, ls)
  - Abstract or paper description
  - Citation / BibTeX blocks
  - API reference lists with no example calls
  - A command only counts if it actually RUNS the tool/model, not just sets up the environment.
    If cd or conda commands appear alongside a real run command, include only the run command.

CONFIDENCE RULES:
  - Use "high"   in almost all cases.
  - Use "medium" ONLY if the content was clearly truncated mid-sentence.
  - Use "low"    ONLY if absolutely no files were fetched.

Respond with a JSON object ONLY — no extra text, no markdown fences:
{
  "has_usage_examples": true | false,
  "confidence": "high" | "medium" | "low",
  "evidence": "<one sentence summarising what you found overall>",
  "example_types": ["distinct", "types", "found"],
  "example_files": [
    {
      "path": "<exact file path as it appears in the ### FILE / ### NOTEBOOK / ### SCRIPT header>",
      "type": "<one of: notebook | example script | README code snippet | CLI usage | tutorial file | quickstart guide | other>",
      "description": "<one sentence: what this specific file demonstrates>",
      "commands": ["<first exact code block or command>", "<second code block if present>"]
    }
  ]
}

IMPORTANT rules for example_files:
  - List EVERY individual example file you found — do not stop at one.
  - For a README that contains multiple distinct usage sections, list it once
    with type "README code snippet".
  - For each notebook or script file, list it as its own entry.
  - If no examples exist, use an empty list [].
  - Paths must be copied EXACTLY from the file headers above (e.g. "README.md",
    "examples/demo.ipynb") — do not invent or guess paths.
  - "commands": list ALL distinct code blocks / commands shown for this file.
    Each entry is one self-contained code block exactly as written in the repo.
    Include BOTH CLI commands (e.g. python run.py ...) AND Python API blocks.
    Use an empty list [] if no commands are shown.
    Do NOT merge multiple blocks into one string.

Return ONLY valid JSON. No explanation, no markdown, no preamble.
"""

RETRY_SYSTEM_PROMPT = """\
You are an expert code reviewer. A previous analysis of this repository returned
uncertain confidence. Re-examine the content and list ALL usage examples found.

Hard rules:
  - Every .ipynb file IS a usage example — list each one.
  - A README code block (``` fences) showing how to call the library → list it.
  - An examples/, tutorials/, or demo/ directory mentioned in the README → list
    the folder as one entry of type "demo folder".
  - A purely abstract / citation README with zero runnable content → empty list.
  - You MUST return "high" confidence. Only "medium" if genuinely mid-truncation.

Respond with a JSON object ONLY:
{
  "has_usage_examples": true | false,
  "confidence": "high",
  "evidence": "<one decisive sentence>",
  "example_types": ["list", "of", "found", "types"],
  "example_files": [
    {"path": "<exact path>", "type": "<type>", "description": "<one sentence>", "commands": ["<code block 1>", "<code block 2 if any>"]}
  ]
}
Return ONLY valid JSON.
"""

TARGETED_README_PROMPT = """\
You are an expert code reviewer. The README was previously truncated; you now
have a larger portion. List ALL usage examples — code blocks, CLI examples,
references to notebooks or example files.

Respond with a JSON object ONLY:
{
  "has_usage_examples": true | false,
  "confidence": "high",
  "evidence": "<one decisive sentence>",
  "example_types": ["list", "of", "found", "types"],
  "example_files": [
    {"path": "<exact path>", "type": "<type>", "description": "<one sentence>", "commands": ["<code block 1>", "<code block 2 if any>"]}
  ]
}
Return ONLY valid JSON.
"""

DEEP_SCAN_PROMPT = """\
You are an expert code reviewer. Additional files (notebooks, scripts, example
folders) have been fetched. List EVERY usage example you find in the content.

Respond with a JSON object ONLY:
{
  "has_usage_examples": true | false,
  "confidence": "high",
  "evidence": "<one decisive sentence>",
  "example_types": ["list", "of", "found", "types"],
  "example_files": [
    {"path": "<exact path>", "type": "<type>", "description": "<one sentence>", "commands": ["<code block 1>", "<code block 2 if any>"]}
  ]
}
Return ONLY valid JSON.
"""


# =============================================================================
# LLM CALL WRAPPER
# =============================================================================

def llm_check_usage(repo_content, paper_title, system_prompt=None):
    """
    Send the combined repo content to the LLM and parse its structured response.

    Returns a dict with:
        has_usage_examples  bool | None
        confidence          "high" | "medium" | "low"
        evidence            str
        example_types       list[str]
        example_files       list[{"path", "type", "description"}]
    """

    def _call(content, token_budget=None):
        user_message = (
            f"Paper: {paper_title}\n\n"
            "Below are the contents of key files from its GitHub repository.\n"
            "Find ALL usage examples present.\n\n"
            + (content if content else "[No relevant files found in the repository]")
        )
        # Scale output budget with input size: large repos can have many examples
        # and the JSON list grows accordingly.  Cap at 4 000 to stay safe.
        if token_budget is None:
            token_budget = 8000 if len(content) > 50_000 else 4000
        response = client.chat.completions.create(
            model=DEPLOYMENT,
            messages=[
                {"role": "system", "content": system_prompt or SYSTEM_PROMPT},
                {"role": "user",   "content": user_message},
            ],
            max_completion_tokens=token_budget,
            response_format={"type": "json_object"},
        )
        TOKEN_USAGE.add_from_response(response)
        return response

    response    = _call(repo_content)
    raw_content = response.choices[0].message.content

    # Empty response = context window overflow → retry with truncated content
    # Also raise the output budget on the retry: the empty response may be caused
    # by truncated JSON being silently discarded, not just by input overflow.
    if not raw_content:
        truncated = repo_content[:30_000] if repo_content else ""
        print(f"\n  [empty response — retrying with {len(truncated):,} chars]",
              end=" ", flush=True)
        response  = _call(truncated, token_budget=16000)
        raw_content = response.choices[0].message.content

    if not raw_content:
        return {
            "has_usage_examples": None,
            "confidence":  "low",
            "evidence":    "Empty LLM response after retry",
            "example_types": [],
            "example_files": [],
        }

    raw_preview = raw_content.strip()
    raw_preview = re.sub(r"^```(?:json)?\s*", "", raw_preview)
    raw_preview = re.sub(r"\s*```$", "", raw_preview)

    print("\nRAW LLM OUTPUT:\n", raw_preview[:1500])

    result = parse_llm_json_response(
        raw_content=raw_content,
        empty_payload={
            "has_usage_examples": None,
            "confidence": "low",
            "evidence": "Empty LLM response after retry",
            "example_types": [],
            "example_files": [],
        },
        required_list_fields=("example_types", "example_files"),
    )

    # Backward compatibility: old prompt variants may return a single string
    # field named example_file instead of the newer example_files list.
    if "example_file" in result and not result.get("example_files"):
        ef = result.pop("example_file")
        result["example_files"] = (
            [{"path": ef, "type": "other", "description": ""}] if ef else []
        )
    else:
        result.pop("example_file", None)

    return result


# =============================================================================
# PER-PAPER ORCHESTRATION
# =============================================================================

def build_example_entries(example_files_raw, owner, repo):
    """
    Convert the LLM's raw example_files list into enriched dicts, each with a
    GitHub hyperlink added.

    Input  (from LLM): [{"path": "examples/demo.ipynb", "type": "notebook",
                          "description": "Demonstrates the core API"}]
    Output:            [{"path": ..., "type": ..., "description": ...,
                          "link": "https://github.com/…/blob/HEAD/examples/demo.ipynb"}]
    """
    entries = []
    for item in (example_files_raw or []):
        path = (item.get("path") or "").strip()
        if not path:
            continue
        cmds = item.get("commands", [])
        if isinstance(cmds, str):   # graceful fallback if LLM returns a string
            cmds = [cmds] if cmds else []
        entry = {
            "path":        path,
            "type":        item.get("type", "other"),
            "description": item.get("description", ""),
            "commands":    cmds,
            "link":        f"https://github.com/{owner}/{repo}/blob/HEAD/{path}",
        }
        entries.append(entry)
    return entries


def check_paper(paper):
    """
    Full pipeline for one paper:
        1. Validate URL (GitHub only)
        2. Fetch repo files
        3. Ask LLM to list ALL examples
        4. Retry if uncertain
        5. Build GitHub links for every example found

    Returns (result_dict, repo_content_string).
    """
    url = paper.get("repo", "")

    result = {
        "title":           paper["title"],
        "repo":            url,
        "status":          None,   # "yes" | "no" | "skipped" | "error"
        "confidence":      None,
        "evidence":        "",
        "example_types":   [],
        "example_entries": [],     # list of {path, type, description, link}
        "files_checked":   [],
        "all_example_meta": [],    # all example-like files spotted in the repo tree
        "note":            "",
    }

    # ── Non-GitHub repos: record clearly rather than silently skip ────────────
    if not is_github(url):
        result["status"] = "skipped"
        result["note"]   = (
            f"Not a GitHub repo — manual review needed. URL: {url}"
            if url else "No repo URL provided"
        )
        return result, ""

    owner, repo = parse_github_repo(url)
    if not owner:
        result["status"] = "error"
        result["note"]   = "Could not parse GitHub URL"
        return result, ""

    try:
        repo_content, files_checked, all_example_meta = collect_repo_content(owner, repo)
        result["files_checked"]    = files_checked
        result["all_example_meta"] = all_example_meta

        llm_result = llm_check_usage(repo_content, paper["title"])

        # Auto-retry if uncertain
        if llm_result.get("confidence") in ("medium", "low"):
            print(f"\n  [retry — confidence was {llm_result.get('confidence')}]",
                  end=" ", flush=True)
            time.sleep(0.5)
            retry_content = repo_content[:80_000] if len(repo_content) > 80_000 else repo_content
            retry = llm_check_usage(retry_content, paper["title"],
                                    system_prompt=RETRY_SYSTEM_PROMPT)
            if retry.get("has_usage_examples") is not None:
                llm_result = retry

        result["confidence"]    = llm_result.get("confidence", "unknown")
        result["evidence"]      = llm_result.get("evidence", "")
        result["example_types"] = llm_result.get("example_types", [])

        # Build enriched example entries (path + type + description + link)
        result["example_entries"] = build_example_entries(
            llm_result.get("example_files", []), owner, repo
        )

        if llm_result.get("has_usage_examples") is None:
            result["status"]     = "error"
            result["note"]       = "Empty LLM response — content may be too large for this model"
            result["confidence"] = "low"
        else:
            result["status"] = "yes" if llm_result["has_usage_examples"] else "no"

    except RuntimeError as e:
        result["status"] = "error"
        result["note"]   = str(e)
        return result, ""
    except Exception as e:
        result["status"] = "error"
        result["note"]   = f"Unexpected error: {e}"
        return result, ""

    return result, repo_content


# =============================================================================
# GROUND-TRUTH EVALUATION
# =============================================================================

def _normalise_ground_truth_label(value):
    """Map common ground-truth label formats to a boolean, or None if unknown."""
    if isinstance(value, bool):
        return value

    if value is None:
        return None

    text = str(value).strip().lower()
    positive = {
        "yes", "true", "1", "has_usage_examples", "has examples",
        "present", "exists", "with_examples", "with examples",
    }
    negative = {
        "no", "false", "0", "missing", "none", "absent",
        "without_examples", "without examples", "no examples",
    }

    if text in positive:
        return True
    if text in negative:
        return False
    return None


def evaluate_against_ground_truth(results):
    """Compute binary metrics when PAPERS includes a usable ground_truth field."""
    gt_by_title = {}
    for paper in PAPERS:
        gt = _normalise_ground_truth_label(paper.get("ground_truth"))
        if gt is not None:
            gt_by_title[paper.get("title", "")] = gt

    per_paper = []
    true_pos = false_pos = true_neg = false_neg = 0
    evaluated = 0

    for row in results:
        title = row.get("title", "")
        if title not in gt_by_title:
            continue

        gt_bool = gt_by_title[title]
        status = (row.get("status") or "").lower()
        if status == "yes":
            pred_bool = True
        elif status == "no":
            pred_bool = False
        else:
            pred_bool = None

        if pred_bool is None:
            match = "unknown"
            note = f"prediction unavailable (status={status or 'n/a'})"
            if row.get("note"):
                note += f"; {row['note']}"
        else:
            evaluated += 1
            if pred_bool and gt_bool:
                true_pos += 1
            elif pred_bool and not gt_bool:
                false_pos += 1
            elif not pred_bool and not gt_bool:
                true_neg += 1
            else:
                false_neg += 1

            match = "correct" if pred_bool == gt_bool else "wrong"
            note = row.get("note") or ""

        per_paper.append({
            "title": title,
            "prediction": "yes" if pred_bool is True else "no" if pred_bool is False else "unknown",
            "ground_truth": "yes" if gt_bool else "no",
            "match": match,
            "note": note,
        })

    labelled = len(per_paper)
    accuracy = ((true_pos + true_neg) / evaluated) if evaluated else 0.0
    precision = (true_pos / (true_pos + false_pos)) if (true_pos + false_pos) else 0.0
    recall = (true_pos / (true_pos + false_neg)) if (true_pos + false_neg) else 0.0
    f1 = (2 * precision * recall / (precision + recall)) if (precision + recall) else 0.0

    return {
        "labelled": labelled,
        "evaluated": evaluated,
        "accuracy": accuracy,
        "precision": precision,
        "recall": recall,
        "f1": f1,
        "true_pos": true_pos,
        "false_pos": false_pos,
        "true_neg": true_neg,
        "false_neg": false_neg,
        "per_paper": per_paper,
    }


def print_ground_truth_report(metrics):
    """Print a compact summary of ground-truth comparison metrics."""
    if metrics.get("labelled", 0) == 0:
        print("\nGround-truth evaluation skipped: no usable ground_truth labels in PAPERS.")
        return

    print("\n" + "=" * 78)
    print("GROUND-TRUTH REPORT")
    print("=" * 78)
    print(f"Labelled papers : {metrics['labelled']}")
    print(f"Evaluated       : {metrics['evaluated']}")
    print(f"Accuracy        : {metrics['accuracy']:.1%}")
    print(f"Precision       : {metrics['precision']:.1%}")
    print(f"Recall          : {metrics['recall']:.1%}")
    print(f"F1 score        : {metrics['f1']:.3f}")
    print(
        "Confusion matrix: "
        f"TP={metrics['true_pos']}, FP={metrics['false_pos']}, "
        f"TN={metrics['true_neg']}, FN={metrics['false_neg']}"
    )

# =============================================================================
# CONFIDENCE DIAGNOSIS & SELF-HEALING
# =============================================================================

DIAGNOSIS_RULES = [
    {
        "id":    "content_truncated",
        "label": "Content was truncated (hit MAX_CONTENT_CHARS limit)",
        "fix":   "Re-fetch README with a larger budget",
        "check": lambda r, content: len(content) >= MAX_CONTENT_CHARS - 100,
    },
    {
        "id":    "notebooks_not_fetched",
        "label": "Notebooks exist but were not fully fetched (char budget exhausted)",
        "fix":   "Fetch notebook content and re-analyse",
        "check": lambda r, content: (
            any(m["kind"] == "notebook" for m in r.get("all_example_meta", []))
            and not any(
                f.lower().endswith((".ipynb", ".rmd", ".qmd"))
                for f in (r.get("files_checked") or [])
            )
        ),
    },
    {
        "id":    "no_example_files",
        "label": "No example / tutorial / demo files detected",
        "fix":   "Scan examples/, tutorials/, and demo/ folders more deeply",
        "check": lambda r, content: (
            not any(
                kw in " ".join(r.get("files_checked") or []).lower()
                for kw in ["example", "tutorial", "demo", "notebook", "quickstart"]
            )
        ),
    },
    {
        "id":    "readme_mentions_examples",
        "label": "README references example files/folders that were not fetched",
        "fix":   "Fetch the referenced files and re-analyse",
        "check": lambda r, content: bool(
            re.search(
                r"(example[s]?|tutorial[s]?|demo|notebook|colab|quickstart)",
                content, re.I
            )
            and r.get("status") == "no"
        ),
    },
]


def diagnose_result(result, repo_content):
    return diagnose_with_rules(result, repo_content, DIAGNOSIS_RULES)


def heal_result(result, repo_content, owner, repo):
    """
    Fetch additional content targeted at the diagnosed problem and re-run the LLM.
    Returns an updated result dict, or the original if healing failed.
    """
    diagnoses = diagnose_result(result, repo_content)
    primary   = diagnoses[0]["id"]

    healed_content = repo_content
    healed_prompt  = RETRY_SYSTEM_PROMPT

    if primary == "content_truncated":
        readme_path = next(
            (f for f in (result.get("files_checked") or []) if "readme" in f.lower()),
            None,
        )
        if readme_path:
            big = fetch_file_content(owner, repo, readme_path)
            if big:
                # Split into 15k chunks and query each, then merge all found examples
                chunk_size = 30_000
                chunks = [big[i:i+chunk_size] for i in range(0, min(len(big), 200_000), chunk_size)]
                all_example_files = []
                last_result = None
                for chunk_idx, chunk in enumerate(chunks):
                    chunk_content = f"### FILE: {readme_path} (part {chunk_idx+1}/{len(chunks)})\n{chunk}"
                    chunk_result  = llm_check_usage(chunk_content, result["title"],
                                                    system_prompt=TARGETED_README_PROMPT)
                    if chunk_result.get("has_usage_examples"):
                        all_example_files.extend(chunk_result.get("example_files", []))
                    last_result = chunk_result
                    time.sleep(0.3)

                if last_result:
                    # Deduplicate by path
                    seen  = set()
                    deduped = []
                    for ef in all_example_files:
                        if ef.get("path") not in seen:
                            seen.add(ef.get("path"))
                            deduped.append(ef)
                    last_result["example_files"]      = deduped
                    last_result["has_usage_examples"] = len(deduped) > 0
                    last_result["confidence"]         = "high"
                    healed_content = big[:60_000]   # keep for reference
                    llm_result     = last_result
                    healed = dict(result)
                    healed["confidence"]      = "high"
                    healed["status"]          = "yes" if last_result["has_usage_examples"] else "no"
                    healed["evidence"]        = last_result.get("evidence", "")
                    healed["example_types"]   = last_result.get("example_types", [])
                    healed["example_entries"] = build_example_entries(deduped, owner, repo)
                    healed["note"]            = "[auto-healed: chunked README scan]"
                    return healed

    elif primary in ("notebooks_not_fetched", "no_example_files", "readme_mentions_examples"):
        all_files   = list_all_repo_files(owner, repo)
        extra_paths = []
        for f in all_files:
            name = f["path"].lower()
            ext  = "." + name.rsplit(".", 1)[-1] if "." in name else ""
            if ext in USAGE_EXTENSIONS:
                extra_paths.append(f["path"])
            elif ext in SCRIPT_EXTENSIONS and any(
                kw in name for kw in
                ["example", "demo", "tutorial", "quickstart", "notebook", "usage", "sample"]
            ):
                extra_paths.append(f["path"])
            elif any(name.startswith(p) for p in EXAMPLE_FOLDER_PREFIXES):
                extra_paths.append(f["path"])

        extra_paths.sort(key=lambda p: (p.count("/"), p.lower()))
        extra_paths = extra_paths[:30]   # fetch up to 15 additional files

        extras = []
        for p in extra_paths:
            raw = fetch_file_content(owner, repo, p)
            if raw:
                if p.lower().endswith(".ipynb"):
                    snippet = summarise_notebook(raw, max_chars=15_000)
                else:
                    snippet = raw[:15_000]
                extras.append(f"### FILE: {p}\n{snippet}")
                time.sleep(0.15)

        if extras:
            healed_content = healed_content + "\n\n" + "\n\n".join(extras)
            healed_prompt  = DEEP_SCAN_PROMPT

    llm_result = llm_check_usage(healed_content, result["title"],
                                  system_prompt=healed_prompt)
    if llm_result.get("has_usage_examples") is None:
        return result

    healed = dict(result)
    healed["confidence"]    = llm_result.get("confidence", "high")
    healed["status"]        = "yes" if llm_result["has_usage_examples"] else "no"
    healed["evidence"]      = llm_result.get("evidence", "")
    healed["example_types"] = llm_result.get("example_types", [])
    healed["example_entries"] = build_example_entries(
        llm_result.get("example_files", []), owner, repo
    )
    healed["note"] = f"[auto-healed: {primary}] " + (result.get("note") or "")
    return healed


# =============================================================================
# CONSOLE REPORTING
# =============================================================================

def print_results(results):
    W = 110
    print("\n" + "=" * W)
    print(f"{'#':<4} {'STATUS':<10} {'CONF':<8} {'# EX':<6} {'EXAMPLE TYPES':<30} TITLE")
    print("=" * W)

    for i, r in enumerate(results, 1):
        icon = {
            "yes": "YES", "no": "NO", "skipped": "SKIP", "error": "ERR",
        }.get(r["status"], "?")

        types_str = ", ".join(r.get("example_types", []))[:28]
        if not types_str and r.get("note"):
            types_str = r["note"][:28]

        n_ex        = len(r.get("example_entries", []))
        title_short = r["title"][:48] + ("..." if len(r["title"]) > 48 else "")
        conf        = r.get("confidence") or "-"
        print(f"{i:<4} {icon:<10} {conf:<8} {n_ex:<6} {types_str:<30} {title_short}")

        if r.get("evidence"):
            print(f"       evidence: {r['evidence']}")
        for ex in r.get("example_entries", []):
            print(f"       [{ex['type']}] {ex['path']}  →  {ex['description']}")

    print("=" * W)
    counts  = count_statuses(results)
    yes     = counts.get("yes", 0)
    no      = counts.get("no", 0)
    skipped = counts.get("skipped", 0)
    errors  = counts.get("error", 0)
    total_ex = sum(len(r.get("example_entries", [])) for r in results)
    print(
        f"\nSUMMARY: {yes} repos have examples | {no} missing | "
        f"{skipped} skipped | {errors} errors | {total_ex} total example files found\n"
    )


# =============================================================================
# EXCEL OUTPUT
# =============================================================================

def save_results(results, path=None):
    """
    Write two Excel sheets:

    Sheet 1 — "Results"  (one row per paper)
        Columns: #, Status, Confidence, Title, Repo, # Examples, Example Types,
                 Evidence, All Example Links (newline-separated), Files Checked, Note

    Sheet 2 — "All Examples"  (one row per individual example file)
        Columns: Paper #, Paper Title, Repo, File Path, Type, Description, Link

    Sheet 3 — "Skipped / Errors"  (one row per paper that was skipped or errored)
        Columns: #, Status, Title, Repo / URL, Note

    Sheet 4 — "Summary"  (totals + type breakdown)
    """
    if path is None:
        path = Path(__file__).resolve().parent / "usage_examples_results.xlsx"

    wb = Workbook()

    # ── Shared style helpers ───────────────────────────────────────────────────
    def hdr_cell(ws, row, col, value, fill_hex="1F5C99"):
        cell            = ws.cell(row=row, column=col, value=value)
        cell.font       = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill       = PatternFill("solid", start_color=fill_hex)
        cell.alignment  = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
        cell.border     = _border()
        return cell

    def _border():
        t = Side(style="thin", color="CCCCCC")
        return Border(left=t, right=t, top=t, bottom=t)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap   = Alignment(horizontal="left",   vertical="top",    wrap_text=True)

    # =========================================================================
    # SHEET 1 — Results (one row per paper)
    # =========================================================================
    ws1        = wb.active
    ws1.title  = "Results"
    hdrs1      = ["#", "Status", "Confidence", "Title", "Repo",
                  "# Examples", "Example Types", "Evidence",
                  "All Example Links", "Files Checked", "Note"]
    widths1    = [5, 10, 12, 45, 40, 10, 35, 55, 60, 50, 35]

    for col, (h, w) in enumerate(zip(hdrs1, widths1), 1):
        hdr_cell(ws1, 1, col, h)
        ws1.column_dimensions[get_column_letter(col)].width = w
    ws1.row_dimensions[1].height = 20

    for row_idx, r in enumerate(results, 2):
        # Build a newline-separated list of all example links for the cell
        all_links_text = "\n".join(
            f"{ex['path']}  ({ex['type']})"
            for ex in r.get("example_entries", [])
        )

        row_data = [
            row_idx - 1,
            (r.get("status") or "").upper(),
            r.get("confidence") or "",
            r.get("title") or "",
            r.get("repo") or "",
            len(r.get("example_entries", [])),
            ", ".join(r.get("example_types") or []),
            r.get("evidence") or "",
            all_links_text,
            ", ".join(r.get("files_checked") or []),
            r.get("note") or "",
        ]
        fill_color = STATUS_FILL_COLORS.get(r.get("status"), "FFFFFF")
        row_fill   = PatternFill("solid", start_color=fill_color)

        for col_idx, value in enumerate(row_data, 1):
            cell           = ws1.cell(row=row_idx, column=col_idx, value=value)
            cell.font      = Font(name="Arial", size=10)
            cell.border    = _border()
            cell.alignment = center if col_idx <= 3 else wrap
            if col_idx == 2:   # colour the Status cell
                cell.fill = row_fill

        # Dynamic height: account for both newlines and wrapped long text
        def estimate_lines(value, col_width=60):
            text = str(value)
            lines = text.split("\n")
            total = sum(max(1, len(line) // col_width + 1) for line in lines)
            return total

        max_lines = max(estimate_lines(v) for v in row_data)
        ws1.row_dimensions[row_idx].height = max(20, max_lines * 15 * 1.3)

    # =========================================================================
    # SHEET 2 — All Examples (one row per individual example file)
    # =========================================================================
    ws2       = wb.create_sheet("All Examples")
    hdrs2     = ["Paper #", "Paper Title", "Repo", "File Path",
                 "Type", "Description", "Example Command", "GitHub Link"]
    widths2   = [8, 50, 45, 55, 22, 60, 70, 70]

    for col, (h, w) in enumerate(zip(hdrs2, widths2), 1):
        hdr_cell(ws2, 1, col, h)
        ws2.column_dimensions[get_column_letter(col)].width = w
    ws2.row_dimensions[1].height = 20

    ex_row = 2
    for paper_num, r in enumerate(results, 1):
        for ex in r.get("example_entries", []):
            row_vals = [
                paper_num,
                r.get("title") or "",
                r.get("repo") or "",
                ex.get("path") or "",
                ex.get("type") or "",
                ex.get("description") or "",
                "\n---\n".join(ex.get("commands") or []),
                ex.get("link") or "",
            ]
            for col_idx, value in enumerate(row_vals, 1):
                is_link = (col_idx == 8)
                cell    = ws2.cell(row=ex_row, column=col_idx, value=value)
                cell.font      = Font(name="Arial", size=10)
                cell.border    = _border()
                cell.alignment = center if col_idx == 1 else wrap
                if is_link and value:
                    cell.hyperlink = value
                    cell.font = Font(name="Arial", size=10,
                                     color="0563C1", underline="single")
            def estimate_lines(value, col_width=60):
                text = str(value)
                lines = text.split("\n")
                total = sum(max(1, len(line) // col_width + 1) for line in lines)
                return total

            max_lines = max(estimate_lines(v) for v in row_vals)
            ws2.row_dimensions[ex_row].height = max(20, max_lines * 15 * 1.3)
            ex_row += 1  # ← add this line

    # =========================================================================
    # SHEET 3 — Skipped / Errors
    # =========================================================================
    ws3      = wb.create_sheet("Skipped & Errors")
    hdrs3    = ["#", "Status", "Title", "Repo / URL", "Note"]
    widths3  = [5, 10, 55, 55, 60]

    for col, (h, w) in enumerate(zip(hdrs3, widths3), 1):
        hdr_cell(ws3, 1, col, h, fill_hex="7B3F00")  # brown header to stand out
        ws3.column_dimensions[get_column_letter(col)].width = w
    ws3.row_dimensions[1].height = 20

    se_row = 2
    for paper_num, r in enumerate(results, 1):
        if r.get("status") not in ("skipped", "error"):
            continue
        fill_color = STATUS_FILL_COLORS.get(r.get("status"), "FFFFFF")
        row_fill   = PatternFill("solid", start_color=fill_color)
        row_vals   = [
            paper_num,
            (r.get("status") or "").upper(),
            r.get("title") or "",
            r.get("repo") or "",
            r.get("note") or "",
        ]
        for col_idx, value in enumerate(row_vals, 1):
            cell           = ws3.cell(row=se_row, column=col_idx, value=value)
            cell.font      = Font(name="Arial", size=10)
            cell.border    = _border()
            cell.alignment = center if col_idx <= 2 else wrap
            if col_idx == 2:
                cell.fill = row_fill
        se_row += 1

    # =========================================================================
    # SHEET 4 — Summary
    # =========================================================================
    ws4      = wb.create_sheet("Summary")
    ws4.column_dimensions["A"].width = 42
    ws4.column_dimensions["B"].width = 15

    hdr_cell(ws4, 1, 1, "Metric")
    hdr_cell(ws4, 1, 2, "Value")
    ws4.row_dimensions[1].height = 20

    counts   = count_statuses(results)
    yes      = counts.get("yes", 0)
    no       = counts.get("no", 0)
    skipped  = counts.get("skipped", 0)
    errors   = counts.get("error", 0)
    total    = len(results)
    total_ex = sum(len(r.get("example_entries", [])) for r in results)

    type_counter = Counter()
    for r in results:
        for t in (r.get("example_types") or []):
            type_counter[t] += 1

    summary_rows = [
        ("Total Repos Checked",           total),
        ("Repos with Usage Examples",     yes),
        ("Repos Missing Usage Examples",  no),
        ("Skipped (non-GitHub)",          skipped),
        ("Errors",                        errors),
        ("Total Example Files Found",     total_ex),
        ("Avg Examples per Repo (w/ any)",
         f"={total_ex}/{yes if yes else 1}"),
        ("Coverage (%)",
         coverage_formula(yes, total)),
        ("", ""),
        ("Example Type Breakdown", ""),
    ]
    for t, cnt in type_counter.most_common():
        summary_rows.append((f"  • {t}", cnt))

    for r_idx, (label, value) in enumerate(summary_rows, 2):
        ws4.cell(row=r_idx, column=1, value=label).font = Font(name="Arial", size=10)
        ws4.cell(row=r_idx, column=2, value=value).font = Font(name="Arial", size=10)
        ws4.cell(row=r_idx, column=2).alignment = center

    # =========================================================================
    # SHEET 5 — Ground Truth Comparison (only written when labels exist)
    # =========================================================================
    metrics = evaluate_against_ground_truth(results)
    if metrics["labelled"] > 0:
        ws5      = wb.create_sheet("Ground Truth")
        hdrs5    = ["#", "Paper Title", "Prediction", "Ground Truth", "Match", "Note"]
        widths5  = [5, 60, 14, 14, 10, 55]
        for col, (h, w) in enumerate(zip(hdrs5, widths5), 1):
            hdr_cell(ws5, 1, col, h, fill_hex="2E7D32")
            ws5.column_dimensions[get_column_letter(col)].width = w
        ws5.row_dimensions[1].height = 20

        MATCH_COLORS = {"correct": "C6EFCE", "wrong": "FFC7CE"}
        for r_idx, p in enumerate(metrics["per_paper"], 2):
            row_vals = [
                r_idx - 1,
                p["title"],
                p["prediction"],
                p["ground_truth"],
                p["match"],
                p["note"],
            ]
            fill_hex = MATCH_COLORS.get(p["match"], "FFFFFF")
            row_fill = PatternFill("solid", start_color=fill_hex)
            for col_idx, value in enumerate(row_vals, 1):
                cell           = ws5.cell(row=r_idx, column=col_idx, value=value)
                cell.font      = Font(name="Arial", size=10)
                cell.border    = _border()
                cell.alignment = center if col_idx in (1, 3, 4, 5) else wrap
                if col_idx == 5:
                    cell.fill = row_fill

        # Summary block below the table
        summary_start = len(metrics["per_paper"]) + 3
        summary_pairs = [
            ("Labelled papers",  metrics["labelled"]),
            ("Accuracy",         f"{metrics['accuracy']:.1%}"),
            ("Precision",        f"{metrics['precision']:.1%}"),
            ("Recall",           f"{metrics['recall']:.1%}"),
            ("F1 score",         f"{metrics['f1']:.3f}"),
            ("True positives",   metrics["true_pos"]),
            ("False positives",  metrics["false_pos"]),
            ("True negatives",   metrics["true_neg"]),
            ("False negatives",  metrics["false_neg"]),
        ]
        for i, (label, value) in enumerate(summary_pairs, summary_start):
            ws5.cell(row=i, column=1, value=label).font = Font(name="Arial", size=10, bold=True)
            ws5.cell(row=i, column=2, value=value).font = Font(name="Arial", size=10)
            ws5.cell(row=i, column=2).alignment = center

        wb.save(path)
    print(f"\nFull results saved to: {path}")
    print(f"  → Results sheet:        {total} papers")
    print(f"  → All Examples sheet:   {total_ex} individual example files")
    se_count = skipped + errors
    print(f"  → Skipped & Errors:     {se_count} entries")
    metrics = evaluate_against_ground_truth(results)
    if metrics["labelled"] > 0:
        print(f"  → Ground Truth sheet:   {metrics['labelled']} labelled, accuracy {metrics['accuracy']:.1%}")
    else:
        print("  → Ground Truth sheet:   not written (no labels in GROUND_TRUTH)")


# =============================================================================
# MAIN
# =============================================================================

def main():
    print(f"Checking {len(PAPERS)} repos for usage examples using LLM ({DEPLOYMENT})...")
    if not GITHUB_TOKEN:
        print("Tip: Set GITHUB_TOKEN in .env to avoid the 60 req/hr GitHub rate limit.\n")

    results       = []
    repo_contents = {}

    # ── Pass 1: analyse every paper ───────────────────────────────────────────
    for i, paper in enumerate(PAPERS, 1):
        owner, repo_name = parse_github_repo(paper.get("repo", ""))
        label = f"{owner}/{repo_name}" if owner else paper.get("repo", "?")
        print(f"[{i:>2}/{len(PAPERS)}] {label} ...", end=" ", flush=True)

        result, content = check_paper(paper)
        results.append(result)
        repo_contents[paper["title"]] = content

        icon  = {"yes": "OK", "no": "NO", "skipped": "SKIP", "error": "ERR"}.get(
            result["status"], "?"
        )
        n_ex  = len(result.get("example_entries", []))
        extra = result.get("confidence") or result.get("note") or ""
        print(f"{icon}  {n_ex} examples  {extra}")
        time.sleep(0.3)

    # ── Confidence report ─────────────────────────────────────────────────────
    print_shared_confidence_report(results, repo_contents, diagnose_result)

    # ── Pass 2: self-heal medium/low confidence results ───────────────────────
    non_high = [i for i, r in enumerate(results)
                if r.get("confidence") in ("medium", "low")]
    if non_high:
        print(f"\n  Auto-healing {len(non_high)} medium/low confidence result(s)...\n")
        for idx in non_high:
            r = results[idx]
            owner, repo = parse_github_repo(r.get("repo", ""))
            if not owner:
                continue
            title_short = r["title"][:55] + ("…" if len(r["title"]) > 55 else "")
            print(f"  healing: {title_short}", end=" ... ", flush=True)
            healed = heal_result(r, repo_contents.get(r["title"], ""), owner, repo)
            results[idx] = healed
            n_before = len(r.get("example_entries", []))
            n_after  = len(healed.get("example_entries", []))
            print(f"conf: {r.get('confidence')} → {healed.get('confidence')}  "
                  f"examples: {n_before} → {n_after}")
            time.sleep(0.5)

        still_non_high = [r for r in results if r.get("confidence") in ("medium", "low")]
        if still_non_high:
            print(f"\n  {len(still_non_high)} result(s) still not high confidence — manual review needed.")
        else:
            print("\n  All results are now high confidence.")

    print_results(results)

    # ── Ground truth evaluation ───────────────────────────────────────────────
    gt_metrics = evaluate_against_ground_truth(results)
    print_ground_truth_report(gt_metrics)

    print_token_usage_report(TOKEN_USAGE, DEPLOYMENT)

    save_results(results)


if __name__ == "__main__":
    main()