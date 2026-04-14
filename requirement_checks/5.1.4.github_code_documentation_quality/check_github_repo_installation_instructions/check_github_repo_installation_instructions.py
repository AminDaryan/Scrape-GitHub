import os, re, sys, time
from pathlib import Path

# Add parent directories to Python path for local imports
_this = Path(__file__).resolve()
sys.path.insert(0, str(_this.parent.parent))          # 5.1.4.github_code_documentation_quality/
sys.path.insert(0, str(_this.parent.parent.parent))   # requirement_checks/

from openpyxl import Workbook
from common.fetch_and_parse_github_repo import (
    load_dotenv, parse_github_repo, is_github, list_all_repo_files, fetch_file_content,
)
from common.confidence_reporting import diagnose_with_rules
from common.token_usage import TokenUsageTracker
from common.result_status import count_statuses
from common.repo_content_helpers import (
    fetch_paths_with_char_budget,
    first_readme_path,
    path_priority_with_readme_first,
)
from common.llm_helpers import llm_call_parse_retry
from common.checker_pipeline import run_checker_pipeline
from common.excel_output import (
    write_header_row, write_results_data_rows, write_summary_sheet,
    thin_border, auto_row_height,
)

load_dotenv()

from papers_from_database import PAPERS
from openai_client import client, AZURE_OPENAI_DEPLOYMENT
from prompts import (
    SYSTEM_PROMPT,
    RETRY_SYSTEM_PROMPT,
    TARGETED_README_PROMPT,
    IMPLICIT_INSTALL_PROMPT,
)

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

# Max chars sent to LLM per repo
MAX_CONTENT_CHARS = 120_000
TOKEN_USAGE = TokenUsageTracker()

# ─── GitHub API helpers ───────────────────────────────────────────────────────


def collect_repo_content(owner, repo):
    """Fetch and concatenate setup-relevant files from a GitHub repo."""
    all_files = list_all_repo_files(owner, repo)
    if not all_files:
        return "", []

    readme_fallback = first_readme_path(all_files)
    selected_paths = []
    for f in all_files:
        fpath_lower = f["path"].lower()
        basename = fpath_lower.split("/")[-1]
        if fpath_lower in TARGET_FILENAMES or basename in TARGET_FILENAMES:
            selected_paths.append(f["path"])

    # Fallback: grab just the root README if nothing matched
    if not selected_paths and readme_fallback:
        selected_paths.append(readme_fallback)

    # Always put the root-level README first so it gets the most budget,
    # then setup files, then nested READMEs last
    selected_paths.sort(
        key=lambda p: path_priority_with_readme_first(p, nested_readme_last=True)
    )

    blocks, fetched, _ = fetch_paths_with_char_budget(
        owner,
        repo,
        selected_paths,
        MAX_CONTENT_CHARS,
        fetch_content=fetch_file_content,
        pause_seconds=0.2,
        header_label="FILE",
    )

    return "\n\n".join(blocks), fetched


# ─── LLM analysis ─────────────────────────────────────────────────────────────
# Prompt templates are defined in neighboring prompts.py.


EMPTY_INSTALLATION_PAYLOAD = {
    "has_installation_instructions": None,
    "confidence": "low",
    "evidence": "Empty LLM response after retry — content may exceed context window",
    "instruction_types": [],
}


def llm_check_installation(repo_content, paper_title, system_prompt=None):
    """Classify whether the repo content contains installation instructions."""
    def build_msg(content):
        return (
            f"Paper: {paper_title}\n\n"
            "Below are the contents of key files fetched from its GitHub repository.\n"
            "Does this repository contain installation instructions?\n\n"
            + (content if content else "[No relevant files found in the repository]")
        )

    return llm_call_parse_retry(
        client=client,
        deployment=AZURE_OPENAI_DEPLOYMENT,
        system_prompt=system_prompt or SYSTEM_PROMPT,
        build_user_message=build_msg,
        content=repo_content,
        token_usage=TOKEN_USAGE,
        empty_payload=EMPTY_INSTALLATION_PAYLOAD,
        required_list_fields=("instruction_types",),
        max_completion_tokens=1000,
        retry_truncate_chars=15_000,
    )


# ─── Per-paper orchestration ──────────────────────────────────────────────────

def check_paper(paper):
    """Validate, fetch, classify, and optionally retry one paper. Returns (result, content)."""
    url = paper.get("repo", "")
    result = {
        "title":             paper["title"],
        "repo":              url,
        "status":            None,
        "confidence":        None,
        "evidence":          "",
        "instruction_types": [],
        "files_checked":     [],
        "installation_link": "",
        "note":              "",
    }

    if not is_github(url):
        result["status"] = "skipped"
        result["note"]   = "Not a GitHub repo"
        return result, ""

    owner, repo = parse_github_repo(url)
    if not owner:
        result["status"] = "error"
        result["note"]   = "Could not parse GitHub URL"
        return result, ""

    try:
        repo_content, files_checked = collect_repo_content(owner, repo)
        result["files_checked"] = files_checked

        llm_result = llm_check_installation(repo_content, paper["title"])

        # Retry with stricter prompt if confidence is not high
        if llm_result.get("confidence") in ("medium", "low"):
            print(f"\n  [retry — confidence was {llm_result.get('confidence')}]", end=" ", flush=True)
            time.sleep(0.5)
            retry_result = llm_check_installation(repo_content, paper["title"], system_prompt=RETRY_SYSTEM_PROMPT)
            if retry_result.get("has_installation_instructions") is not None:
                llm_result = retry_result

        result["confidence"]        = llm_result.get("confidence", "unknown")
        result["evidence"]          = llm_result.get("evidence", "")
        result["instruction_types"] = llm_result.get("instruction_types", [])

        install_file = llm_result.get("installation_file")
        if install_file and owner:
            result["installation_link"] = f"https://github.com/{owner}/{repo}/blob/HEAD/{install_file}"

        if llm_result.get("has_installation_instructions") is None:
            result["status"] = "error"
        else:
            result["status"] = "yes" if llm_result["has_installation_instructions"] else "no"

    except RuntimeError as e:
        result["status"] = "error"
        result["note"]   = str(e)
        return result, ""
    except Exception as e:
        result["status"] = "error"
        result["note"]   = f"Unexpected error: {e}"
        return result, ""

    return result, repo_content


# ─── Confidence diagnosis & self-healing ──────────────────────────────────────

# Root causes and their fixes
DIAGNOSIS_RULES = [
    {
        "id":    "content_truncated",
        "label": "Content was truncated (hit MAX_CONTENT_CHARS limit)",
        "fix":   "Re-fetch with a higher char limit targeting only the README",
        "check": lambda r, content: len(content) >= MAX_CONTENT_CHARS - 100,
    },
    {
        "id":    "no_setup_files",
        "label": "No setup files found — only README was fetched",
        "fix":   "Do a deeper search for any setup/install file in the repo",
        "check": lambda r, content: (
            len(r.get("files_checked", [])) <= 1
            and not any(
                f in (r.get("files_checked") or [])
                for f in ["requirements.txt", "environment.yml", "setup.py",
                          "pyproject.toml", "Dockerfile"]
            )
        ),
    },
    {
        "id":    "install_section_exists_but_truncated",
        "label": "README has an installation section heading but content was cut off",
        "fix":   "Fetch only the README at double the normal char budget",
        "check": lambda r, content: bool(
            re.search(r"#{1,3}\s*(install|setup|getting.started)", content, re.I)
            and len(content) >= MAX_CONTENT_CHARS - 100
        ),
    },
    {
        "id":    "vague_readme",
        "label": "README mentions prerequisites/dependencies but no explicit commands",
        "fix":   "Ask the LLM to re-evaluate with a stricter definition of 'implicit' instructions",
        "check": lambda r, content: bool(
            re.search(r"(prerequisite|requirement|depend|python\s*[\d.]+|version\s*[\d.]+)",
                      content, re.I)
            and not re.search(r"(pip install|conda|docker|apt-get|brew install|yarn|npm install)",
                              content, re.I)
        ),
    },
]

# Additional prompt variants are defined in neighboring prompts.py.


def diagnose_result(result, repo_content):
    """Identify likely root causes for uncertain LLM confidence."""
    return diagnose_with_rules(result, repo_content, DIAGNOSIS_RULES)


def heal_result(result, repo_content, owner, repo):
    """Re-fetch targeted content and re-run the LLM to improve a low-confidence result."""
    diagnoses = diagnose_result(result, repo_content)
    primary = diagnoses[0]["id"]

    healed_content  = repo_content
    healed_prompt   = RETRY_SYSTEM_PROMPT

    if primary in ("content_truncated", "install_section_exists_but_truncated"):
        # Re-fetch README with a much larger budget
        readme_path = next(
            (f for f in (result.get("files_checked") or [])
             if "readme" in f.lower()),
            None
        )
        if readme_path:
            big_content = fetch_file_content(owner, repo, readme_path)
            if big_content:
                healed_content = f"### FILE: {readme_path}\n{big_content[:80_000]}"
                healed_prompt  = TARGETED_README_PROMPT

    elif primary == "no_setup_files":
        # Search the full tree for any setup/install file we might have missed
        all_files = list_all_repo_files(owner, repo)
        extra_paths = []
        for f in all_files:
            name = f["path"].lower()
            if any(kw in name for kw in
                   ["install", "setup", "require", "environment", "docker",
                    "conda", "makefile", "pyproject"]):
                extra_paths.append(f["path"])
        if extra_paths:
            extras = []
            for p in extra_paths[:5]:
                c = fetch_file_content(owner, repo, p)
                if c:
                    extras.append(f"### FILE: {p}\n{c[:8_000]}")
                    time.sleep(0.2)
            if extras:
                healed_content = healed_content + "\n\n" + "\n\n".join(extras)

    elif primary == "vague_readme":
        healed_prompt = IMPLICIT_INSTALL_PROMPT

    llm_result = llm_check_installation(healed_content, result["title"],
                                        system_prompt=healed_prompt)
    if llm_result.get("has_installation_instructions") is None:
        return result  # healing failed — keep original

    healed = dict(result)
    healed["confidence"]        = llm_result.get("confidence", "high")
    healed["status"]            = "yes" if llm_result["has_installation_instructions"] else "no"
    healed["evidence"]          = llm_result.get("evidence", "")
    healed["instruction_types"] = llm_result.get("instruction_types", [])
    healed["note"]              = f"[auto-healed: {primary}] " + (result.get("note") or "")

    install_file = llm_result.get("installation_file")
    if install_file and owner:
        healed["installation_link"] = f"https://github.com/{owner}/{repo}/blob/HEAD/{install_file}"

    return healed


def print_results(results):
    """Print a compact console table of per-paper results."""
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
    counts  = count_statuses(results)
    yes     = counts.get("yes", 0)
    no      = counts.get("no", 0)
    skipped = counts.get("skipped", 0)
    errors  = counts.get("error", 0)
    print(
        f"\nSUMMARY: {yes} have install instructions | "
        f"{no} missing | {skipped} skipped (non-GitHub) | {errors} errors\n"
    )



def save_results(results, path=None):
    """Export results and summary to an Excel workbook."""
    if path is None:
        path = Path(__file__).resolve().parent / "results/installation_instructions_results.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"

    headers = ["#", "Status", "Confidence", "Title", "Repo",
               "Instruction Types", "Evidence", "Installation Link",
               "Files Checked", "Note"]
    col_widths = [5, 10, 12, 45, 40, 35, 50, 45, 45, 30]
    border = thin_border()

    write_header_row(ws, headers, col_widths, fill_hex="2F5496", border=border)

    def row_data_fn(r, num):
        return [
            num,
            (r.get("status") or "").upper(),
            r.get("confidence") or "",
            r.get("title") or "",
            r.get("repo") or "",
            ", ".join(r.get("instruction_types") or []),
            r.get("evidence") or "",
            r.get("installation_link") or "",
            ", ".join(r.get("files_checked") or []),
            r.get("note") or "",
        ]

    write_results_data_rows(
        ws, results, row_data_fn,
        border=border,
        link_cols={8},
        row_height_fn=lambda vals: auto_row_height(vals),
    )

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    write_summary_sheet(
        ws2, results,
        positive_label="Have Installation Instructions",
        negative_label="Missing Installation Instructions",
        fill_hex="2F5496",
        border=border,
    )

    wb.save(path)
    print(f"Full results saved to {path}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    """Entry point: run the two-pass analysis pipeline."""
    run_checker_pipeline(
        papers=PAPERS,
        check_paper_fn=check_paper,
        diagnose_fn=diagnose_result,
        heal_fn=heal_result,
        print_results_fn=print_results,
        save_results_fn=save_results,
        token_usage=TOKEN_USAGE,
        deployment=AZURE_OPENAI_DEPLOYMENT,
        description="installation instructions",
        github_token=GITHUB_TOKEN,
    )


if __name__ == "__main__":
    main()