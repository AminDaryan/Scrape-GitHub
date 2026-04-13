import os, re, sys, time
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
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
    first_readme_path,
    path_priority_with_readme_first,
)

load_dotenv()

sys.path.append(str(Path(__file__).resolve().parent.parent))
from papers_from_database import PAPERS

sys.path.append(str(Path(__file__).resolve().parent.parent.parent))
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
    """Collect and prioritize repository files that may contain setup guidance.

    This function lists all files in a GitHub repository, filters to known
    setup/documentation filenames, then fetches and concatenates their contents
    into a single payload for LLM analysis. Each chunk is prefixed with a
    ``### FILE: <path>`` header so the model can cite exact evidence.

    The ordering is intentional: root README first, then other root files,
    then nested files. This helps preserve the most informative content when
    the MAX_CONTENT_CHARS budget is reached.

    Args:
        owner (str): GitHub account or organization name.
        repo (str): GitHub repository name.

    Returns:
        tuple[str, list[str]]: Concatenated file content and the list of file
        paths that were actually fetched. Returns ("", []) when the repository
        tree could not be listed.
    """
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


def llm_check_installation(repo_content, paper_title, system_prompt=None):
    """Run LLM classification for installation-instruction detection.

    The function sends repository text to the configured Azure OpenAI model,
    expects a JSON object response, and normalizes/parses it via the shared
    parser. If the model returns an empty payload (often context overflow), it
    retries once with a shorter content slice.

    Args:
        repo_content (str): Concatenated repository snippets with file headers.
        paper_title (str): Paper title used to provide context in the prompt.
        system_prompt (str | None): Optional override prompt for targeted
            re-evaluation flows.

    Returns:
        dict: Parsed LLM result containing boolean decision, confidence,
        evidence text, instruction types, and optional installation file.
    """
    def _call(content):
        """Submit one chat-completions request and record token usage."""
        user_message = (
            f"Paper: {paper_title}\n\n"
            "Below are the contents of key files fetched from its GitHub repository.\n"
            "Does this repository contain installation instructions?\n\n"
            + (content if content else "[No relevant files found in the repository]")
        )
        response = client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT,
            messages=[
                {"role": "system", "content": system_prompt or SYSTEM_PROMPT},
                {"role": "user",   "content": user_message},
            ],
            max_completion_tokens=1000,
            response_format={"type": "json_object"}
        )
        TOKEN_USAGE.add_from_response(response)
        return response

    response = _call(repo_content)
    raw_content = response.choices[0].message.content

    # Empty response usually means context window overflow — shrink and retry once
    if not raw_content:
        truncated = repo_content[:15_000] if repo_content else ""
        print(f"\n  [empty response — retrying with {len(truncated):,} chars]", end=" ", flush=True)
        response = _call(truncated)
        raw_content = response.choices[0].message.content

    if not raw_content:
        return {
            "has_installation_instructions": None,
            "confidence": "low",
            "evidence": "Empty LLM response after retry — content may exceed context window",
            "instruction_types": [],
        }

    raw_preview = raw_content.strip()
    raw_preview = re.sub(r"^```(?:json)?\s*", "", raw_preview)
    raw_preview = re.sub(r"\s*```$", "", raw_preview)

    print("\nRAW LLM OUTPUT:\n", raw_preview[:500])

    return parse_llm_json_response(
        raw_content=raw_content,
        empty_payload={
            "has_installation_instructions": None,
            "confidence": "low",
            "evidence": "Empty LLM response after retry — content may exceed context window",
            "instruction_types": [],
        },
        required_list_fields=("instruction_types",),
    )


# ─── Per-paper orchestration ──────────────────────────────────────────────────

def check_paper(paper):
    """Execute end-to-end installation-instruction checks for one paper record.

    This orchestrator validates the repository URL, fetches candidate files,
    calls the LLM classifier, optionally retries with a stricter prompt when
    confidence is not high, and formats a normalized result dictionary used by
    reporting/export code.

    Args:
        paper (dict): Entry from PAPERS with at least ``title`` and ``repo``.

    Returns:
        tuple[dict, str]: Final result object and the raw fetched repository
        content string used for confidence diagnosis.
    """
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

        result["status"]            = "yes" if llm_result.get("has_installation_instructions") else "no"
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
    """Identify likely root causes for uncertain LLM confidence.

    Args:
        result (dict): One per-paper evaluation record.
        repo_content (str): Full text payload that was sent to the LLM.

    Returns:
        list[dict]: Ordered diagnosis matches from ``DIAGNOSIS_RULES``.
    """
    return diagnose_with_rules(result, repo_content, DIAGNOSIS_RULES)


def heal_result(result, repo_content, owner, repo):
    """
    Attempt targeted remediation when confidence is medium/low.

    The function picks the top diagnosis and applies a focused recovery
    strategy (for example: refetch a larger README, search for missed setup
    files, or use a broader prompt for implicit instructions). It then reruns
    the LLM and merges successful updates into the original result structure.

    Args:
        result (dict): Existing per-paper result to improve.
        repo_content (str): Original fetched content payload.
        owner (str): GitHub owner/org.
        repo (str): GitHub repository name.

    Returns:
        dict: Healed result when re-analysis succeeds, otherwise the original
        result unchanged.
    """
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
    """Render a compact console table of per-paper decisions and evidence.

    Args:
        results (list[dict]): Final normalized result objects.
    """
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
    """Export detailed and summary results to an Excel workbook.

    The workbook contains:
    - ``Results``: one row per repository with status, evidence, and links.
    - ``Summary``: aggregate counts and coverage formula.

    Args:
        results (list[dict]): Final normalized results.
        path (Path | None): Optional output path. Defaults to
            ``installation_instructions_results.xlsx`` next to this script.
    """
    if path is None:
        path = Path(__file__).resolve().parent / "installation_instructions_results.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"

    # Styles
    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="2F5496")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["#", "Status", "Confidence", "Title", "Repo", "Instruction Types", "Evidence", "Installation Link", "Files Checked", "Note"]
    col_widths = [5, 10, 12, 45, 40, 35, 50, 45, 45, 30]

    for col_idx, (header, width) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 20

    for row_idx, r in enumerate(results, 2):
        install_link = r.get("installation_link") or ""
        row_data = [
            row_idx - 1,
            (r.get("status") or "").upper(),
            r.get("confidence") or "",
            r.get("title") or "",
            r.get("repo") or "",
            ", ".join(r.get("instruction_types") or []),
            r.get("evidence") or "",
            install_link,
            ", ".join(r.get("files_checked") or []),
            r.get("note") or "",
        ]
        fill_color = STATUS_FILL_COLORS.get(r.get("status"), "FFFFFF")
        row_fill = PatternFill("solid", start_color=fill_color)

        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = center if col_idx <= 3 else wrap
            if col_idx == 2:
                cell.fill = row_fill
            # Render Installation Link column (col 8) as a clickable hyperlink
            if col_idx == 8 and value:
                cell.hyperlink = value
                cell.font = Font(name="Arial", size=10, color="0563C1", underline="single")

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    counts  = count_statuses(results)
    yes     = counts.get("yes", 0)
    no      = counts.get("no", 0)
    skipped = counts.get("skipped", 0)
    errors  = counts.get("error", 0)
    total   = len(results)

    summary_data = [
        ("Total Repos Checked", total),
        ("Have Installation Instructions", yes),
        ("Missing Installation Instructions", no),
        ("Skipped (non-GitHub)", skipped),
        ("Errors", errors),
        ("Coverage (%)", coverage_formula(yes, total)),
    ]

    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 15
    ws2.cell(row=1, column=1, value="Metric").font = Font(name="Arial", bold=True)
    ws2.cell(row=1, column=2, value="Value").font = Font(name="Arial", bold=True)
    for fill_col in [1, 2]:
        ws2.cell(row=1, column=fill_col).fill = PatternFill("solid", start_color="2F5496")
        ws2.cell(row=1, column=fill_col).font = Font(name="Arial", bold=True, color="FFFFFF")
        ws2.cell(row=1, column=fill_col).alignment = center

    for r_idx, (label, value) in enumerate(summary_data, 2):
        ws2.cell(row=r_idx, column=1, value=label).font = Font(name="Arial", size=10)
        ws2.cell(row=r_idx, column=2, value=value).font = Font(name="Arial", size=10)
        ws2.cell(row=r_idx, column=2).alignment = center

    wb.save(path)
    print(f"Full results saved to {path}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    """Run the full two-pass analysis pipeline over all configured papers.

    Workflow:
    1) Evaluate each repository once.
    2) Print a confidence report and diagnoses.
    3) Auto-heal medium/low confidence cases and re-evaluate.
    4) Print token usage, console summary, and save Excel output.
    """
    print(f"Checking {len(PAPERS)} repos for installation instructions "
          f"using LLM ({AZURE_OPENAI_DEPLOYMENT})...")
    if not GITHUB_TOKEN:
        print("Tip: Set GITHUB_TOKEN in your .env to avoid the 60 req/hr GitHub rate limit.\n")

    results      = []
    repo_contents = {}  # title → raw content fetched, used for diagnosis

    # ── Pass 1: check all papers ──────────────────────────────────────────────
    for i, paper in enumerate(PAPERS, 1):
        owner, repo_name = parse_github_repo(paper.get("repo", ""))
        label = f"{owner}/{repo_name}" if owner else paper["repo"]
        print(f"[{i:>2}/{len(PAPERS)}] {label} ...", end=" ", flush=True)

        result, content = check_paper(paper)
        results.append(result)
        repo_contents[paper["title"]] = content

        icon  = {"yes": "OK", "no": "NO", "skipped": "SKIP", "error": "ERR"}.get(result["status"], "?")
        extra = result.get("confidence") or result.get("note") or ""
        print(f"{icon}  {extra}")

        time.sleep(0.3)

    # ── Confidence report + diagnosis ─────────────────────────────────────────
    print_shared_confidence_report(results, repo_contents, diagnose_result)

    # ── Pass 2: self-heal any remaining medium/low confidence results ─────────
    non_high = [i for i, r in enumerate(results) if r.get("confidence") in ("medium", "low")]
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
            print(f"conf: {r.get('confidence')} → {healed.get('confidence')}")
            time.sleep(0.5)

        # Final confidence report after healing
        still_non_high = [r for r in results if r.get("confidence") in ("medium", "low")]
        if still_non_high:
            print(f"\n  {len(still_non_high)} result(s) still not high confidence after healing.")
            print("  These likely require manual review (repo has no machine-readable setup info).")
        else:
            print("\n  All results are now high confidence.")

    print_token_usage_report(TOKEN_USAGE, AZURE_OPENAI_DEPLOYMENT)

    print_results(results)
    save_results(results)


if __name__ == "__main__":
    main()