"""Inline-comments checker for academic paper GitHub repositories.

Fetches source code files via the GitHub API, then asks the LLM to evaluate
whether the code contains meaningful inline comments. Results export to Excel.
"""

import os
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# Add parent directories to Python path for local imports
_this = Path(__file__).resolve()
sys.path.insert(0, str(_this.parent.parent))          # 5.1.4.code_documentation_quality/
sys.path.insert(0, str(_this.parent.parent.parent))   # requirement_checks/

from common.fetch_and_parse_github_repo import (
    load_dotenv, parse_github_repo, is_github, list_all_repo_files, fetch_file_content,
)
from common.token_usage import TokenUsageTracker
from common.result_status import STATUS_FILL_COLORS, count_statuses
from common.repo_content_helpers import (
    fetch_paths_with_char_budget,
    path_priority_with_readme_first,
)
from common.llm_helpers import llm_call_parse_retry
from common.checker_pipeline import run_checker_pipeline
from common.excel_output import (
    write_header_row, write_results_data_rows, write_summary_sheet,
    thin_border, auto_row_height, alignment_center, alignment_wrap_left,
)
from shared.check_paper_common import check_paper_generic

load_dotenv()

from papers_from_database import PAPERS
from openai_client import client, AZURE_OPENAI_DEPLOYMENT
from prompts import SYSTEM_PROMPT
from config import (
    SOURCE_CODE_EXTENSIONS, SKIP_FOLDER_PREFIXES, SKIP_BASENAMES,
    MAX_SOURCE_FILES, MAX_CONTENT_CHARS, PER_FILE_CHAR_CAP,
)

GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")

TOKEN_USAGE = TokenUsageTracker()


# =============================================================================
# GITHUB API HELPERS
# =============================================================================


def _is_source_file(path_lower, basename):
    """Return True if the file looks like a source code file worth checking."""
    ext = "." + basename.rsplit(".", 1)[-1] if "." in basename else ""
    if ext not in SOURCE_CODE_EXTENSIONS:
        return False
    if any(path_lower.startswith(prefix) for prefix in SKIP_FOLDER_PREFIXES):
        return False
    if basename in SKIP_BASENAMES:
        return False
    return True


def collect_repo_content(owner, repo):
    """Fetch source code files from a GitHub repo for inline-comment analysis."""
    all_files = list_all_repo_files(owner, repo)
    if not all_files:
        return "", []

    source_paths = []
    for f in all_files:
        fpath = f["path"]
        fpath_lower = fpath.lower()
        basename = fpath_lower.split("/")[-1]
        if _is_source_file(fpath_lower, basename):
            source_paths.append(fpath)

    if not source_paths:
        return "", []

    # Prioritize: root-level files first, then shallowest, then alphabetical.
    source_paths.sort(key=lambda p: (p.count("/"), p.lower()))

    # Cap to avoid fetching too many files.
    source_paths = source_paths[:MAX_SOURCE_FILES]

    blocks, fetched, _ = fetch_paths_with_char_budget(
        owner,
        repo,
        source_paths,
        MAX_CONTENT_CHARS,
        fetch_content=fetch_file_content,
        pause_seconds=0.15,
        header_label="SOURCE FILE",
        per_file_char_cap=PER_FILE_CHAR_CAP,
    )

    return "\n\n".join(blocks), fetched


# =============================================================================
# LLM CALL WRAPPER
# =============================================================================

EMPTY_INLINE_COMMENTS_PAYLOAD = {
    "has_inline_comments": None,
    "comment_quality": "none",
    "evidence": "Empty LLM response after retry — content may exceed context window",
    "comment_types": [],
    "files_with_comments": [],
}


def llm_check_inline_comments(repo_content, paper_title):
    """Classify whether the repo code contains meaningful inline comments."""
    def build_msg(content):
        return (
            f"Paper: {paper_title}\n\n"
            "Below are the contents of source code files fetched from its GitHub repository.\n"
            "Evaluate whether the code contains meaningful inline comments.\n\n"
            + (content if content else "[No source code files found in the repository]")
        )

    token_budget = 4000 if len(repo_content) > 50_000 else 2000

    return llm_call_parse_retry(
        client=client,
        deployment=AZURE_OPENAI_DEPLOYMENT,
        system_prompt=SYSTEM_PROMPT,
        build_user_message=build_msg,
        content=repo_content,
        token_usage=TOKEN_USAGE,
        empty_payload=EMPTY_INLINE_COMMENTS_PAYLOAD,
        required_list_fields=("comment_types", "files_with_comments"),
        max_completion_tokens=token_budget,
        retry_truncate_chars=30_000,
        retry_max_tokens=8000,
        preview_chars=1000,
    )


# =============================================================================
# PER-PAPER ORCHESTRATION
# =============================================================================

def _collect_wrapper(owner, repo, result):
    """Adapter: call collect_repo_content and return (content, files_checked)."""
    return collect_repo_content(owner, repo)


def _map_inline_comments(result, llm_result, owner, repo):
    """Copy inline-comment-specific fields from LLM result into the paper result."""
    result["comment_quality"]     = llm_result.get("comment_quality", "none")
    result["comment_types"]       = llm_result.get("comment_types", [])
    result["files_with_comments"] = llm_result.get("files_with_comments", [])


def check_paper(paper):
    """Validate, fetch, and classify one paper. Returns a result dict."""
    return check_paper_generic(
        paper,
        extra_defaults={
            "comment_quality":     "none",
            "comment_types":       [],
            "files_with_comments": [],
        },
        collect_content_fn=_collect_wrapper,
        llm_check_fn=llm_check_inline_comments,
        map_llm_result_fn=_map_inline_comments,
        boolean_key="has_inline_comments",
        require_files=True,
        no_files_message="No source code files found in the repository",
    )


# =============================================================================
# CONSOLE REPORTING
# =============================================================================

def print_results(results):
    """Print a compact console table of per-paper results."""
    W = 110
    print("\n" + "=" * W)
    print(f"{'#':<4} {'STATUS':<10} {'QUALITY':<10} {'COMMENT TYPES':<36} TITLE")
    print("=" * W)

    for i, r in enumerate(results, 1):
        icon = {
            "yes": "YES", "no": "NO", "skipped": "SKIP", "error": "ERR",
        }.get(r["status"], "?")

        quality = r.get("comment_quality", "-") or "-"
        types_str = ", ".join(r.get("comment_types", []))[:34]
        if not types_str and r.get("note"):
            types_str = r["note"][:34]

        title_short = r["title"][:46] + ("..." if len(r["title"]) > 46 else "")
        print(f"{i:<4} {icon:<10} {quality:<10} {types_str:<36} {title_short}")

        if r.get("evidence"):
            print(f"       evidence: {r['evidence']}")
        if r.get("files_checked"):
            print(f"       files: {', '.join(r['files_checked'][:5])}"
                  + (f" (+{len(r['files_checked'])-5} more)" if len(r["files_checked"]) > 5 else ""))

    print("=" * W)
    counts  = count_statuses(results)
    yes     = counts.get("yes", 0)
    no      = counts.get("no", 0)
    skipped = counts.get("skipped", 0)
    errors  = counts.get("error", 0)
    print(
        f"\nSUMMARY: {yes} repos have inline comments | "
        f"{no} missing | {skipped} skipped (non-GitHub) | {errors} errors\n"
    )


# =============================================================================
# EXCEL OUTPUT
# =============================================================================

def save_results(results, path=None):
    """Write Excel sheets: Results, Per-File Detail, Summary."""
    if path is None:
        path = Path(__file__).resolve().parent / "results/inline_comments_results.xlsx"

    wb = Workbook()
    border = thin_border()
    center = alignment_center()
    wrap = alignment_wrap_left()

    # ── SHEET 1: Results (one row per paper) ─────────────────────────────────
    ws1 = wb.active
    ws1.title = "Results"
    hdrs1 = ["#", "Status", "Quality", "Title", "Repo",
             "Comment Types", "Evidence", "Files Checked", "Note"]
    widths1 = [5, 10, 10, 45, 40, 35, 55, 50, 35]
    write_header_row(ws1, hdrs1, widths1, fill_hex="2F5496", border=border)

    def results_row_data(r, num):
        return [
            num,
            (r.get("status") or "").upper(),
            r.get("comment_quality") or "none",
            r.get("title") or "",
            r.get("repo") or "",
            ", ".join(r.get("comment_types") or []),
            r.get("evidence") or "",
            ", ".join(r.get("files_checked") or []),
            r.get("note") or "",
        ]

    write_results_data_rows(
        ws1, results, results_row_data,
        border=border,
        row_height_fn=lambda vals: auto_row_height(vals),
    )

    # ── SHEET 2: Per-File Detail (one row per file with comments) ────────────
    ws2 = wb.create_sheet("Per-File Detail")
    hdrs2 = ["Paper #", "Paper Title", "Repo", "File Path",
             "Quality", "Description", "GitHub Link"]
    widths2 = [8, 50, 45, 55, 12, 60, 70]
    write_header_row(ws2, hdrs2, widths2, fill_hex="2F5496", border=border)

    detail_row = 2
    for paper_num, r in enumerate(results, 1):
        owner, repo_name = parse_github_repo(r.get("repo", ""))
        for fc in r.get("files_with_comments", []):
            fpath = (fc.get("path") or "").strip()
            if not fpath:
                continue
            link = (
                f"https://github.com/{owner}/{repo_name}/blob/HEAD/{fpath}"
                if owner else ""
            )
            row_vals = [
                paper_num,
                r.get("title") or "",
                r.get("repo") or "",
                fpath,
                fc.get("quality") or "",
                fc.get("description") or "",
                link,
            ]
            for col_idx, value in enumerate(row_vals, 1):
                cell = ws2.cell(row=detail_row, column=col_idx, value=value)
                cell.font = Font(name="Arial", size=10)
                cell.border = border
                cell.alignment = center if col_idx == 1 else wrap
                if col_idx == 7 and value:
                    cell.hyperlink = value
                    cell.font = Font(name="Arial", size=10,
                                     color="0563C1", underline="single")
            ws2.row_dimensions[detail_row].height = auto_row_height(
                row_vals, line_height=15 * 1.3,
            )
            detail_row += 1

    # ── SHEET 3: Summary ─────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Summary")
    write_summary_sheet(
        ws3, results,
        positive_label="Have Inline Comments",
        negative_label="Missing Inline Comments",
        fill_hex="2F5496",
        border=border,
    )

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"Full results saved to {path}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Entry point: run the inline-comments analysis pipeline."""
    run_checker_pipeline(
        papers=PAPERS,
        check_paper_fn=check_paper,
        print_results_fn=print_results,
        save_results_fn=save_results,
        token_usage=TOKEN_USAGE,
        deployment=AZURE_OPENAI_DEPLOYMENT,
        description="inline comments",
        github_token=GITHUB_TOKEN,
    )


if __name__ == "__main__":
    main()
