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

load_dotenv()

sys.path.append(str(Path(__file__).resolve().parent.parent))
from papers_from_database import PAPERS

sys.path.append(str(Path(__file__).resolve().parent.parent.parent))
from openai_client import client, AZURE_OPENAI_DEPLOYMENT

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
    """
    Fetch all targeted files from the repo and concatenate into a single
    string for the LLM. Returns (content_string, list_of_fetched_paths).
    Root README is always fetched first and given the most budget.
    """
    all_files = list_all_repo_files(owner, repo)
    if not all_files:
        return "", []

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

    # Always put the root-level README first so it gets the most budget,
    # then setup files, then nested READMEs last
    def sort_key(p):
        pl = p.lower()
        if pl in ("readme.md", "readme.rst", "readme.txt", "readme"):
            return 0   # root README first
        if "/" not in pl:
            return 1   # other root-level files second
        if "readme" in pl:
            return 3   # nested READMEs last
        return 2
    selected_paths.sort(key=sort_key)

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
        time.sleep(0.2)

    return "\n\n".join(combined), fetched


# ─── LLM analysis ─────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are an expert code reviewer analysing GitHub repositories for academic papers.

Your task: determine whether the repository contains INSTALLATION INSTRUCTIONS —
i.e. any content that explains how to set up the environment, install dependencies,
or get the code running (pip install, conda, docker, apt-get, virtualenv, etc.).

You will receive the concatenated contents of key files (README, setup files,
requirements, Dockerfiles, Makefiles, etc.).

CONFIDENCE RULES — you must use "high" in almost all cases:
- Use "high" if the files clearly contain installation steps (commands, package lists, environment setup).
- Use "high" if the files are clearly just a paper summary, dataset description, or survey with zero setup content.
- Use "high" if a requirements.txt or environment.yml is present — that IS an installation instruction.
- Use "medium" ONLY if the content was visibly truncated mid-sentence and you genuinely cannot tell.
- Use "low" ONLY if no files at all were fetched.
- A README that only mentions prerequisites by name (no commands) counts as "other" instruction type — still mark true with high confidence.

Respond with a JSON object ONLY — no extra text, no markdown fences:
{
  "has_installation_instructions": true | false,
  "confidence": "high" | "medium" | "low",
  "evidence": "<one sentence describing what you found or why you concluded no instructions exist>",
  "instruction_types": ["list", "of", "found", "types"],
  "installation_file": "<the exact file path where installation instructions were found, e.g. README.md or docs/installation.md, or null if none found>"
}

Possible instruction_types values (use as many as apply):
  pip install, conda environment, docker setup, system packages,
  virtual environment, requirements file, makefile, manual build steps, other
If none found, use an empty list [].
For installation_file: return the file path string exactly as it appears in the "### FILE: <path>" headers above (e.g. "README.md", "docs/installation.md"). Return null if no instructions found.
Return ONLY valid JSON.
Do NOT include any explanation, text, or markdown.
Your entire response must be a single valid JSON object.
"""

RETRY_SYSTEM_PROMPT = """\
You are an expert code reviewer. A previous analysis of this GitHub repository returned uncertain confidence.
Re-examine the content carefully and make a DECISIVE, HIGH-CONFIDENCE determination.

Rules:
- If ANY setup file (requirements.txt, environment.yml, Dockerfile, setup.py, pyproject.toml) is present → true, high confidence, "requirements file" type.
- If the README lists packages or versions to install, even informally → true, high confidence, "other" type.
- If the README is purely a paper abstract, survey, or dataset description with zero setup content → false, high confidence.
- You MUST return "high" confidence. Only return "medium" if the content is genuinely mid-truncation.

Respond with a JSON object ONLY:
{
  "has_installation_instructions": true | false,
  "confidence": "high",
  "evidence": "<one decisive sentence>",
  "instruction_types": ["list", "of", "found", "types"],
  "installation_file": "<file path or null>"
}
Return ONLY valid JSON.
"""


def llm_check_installation(repo_content, paper_title, system_prompt=None):
    """Send repo content to the LLM and parse its structured response.
    If the response is empty (context window overflow), shrink content and retry once.
    """
    def _call(content):
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

# Targeted prompt for repos where only the README section was truncated
TARGETED_README_PROMPT = """\
You are an expert code reviewer. The README for this repo was previously truncated.
You are now receiving a larger portion. Focus specifically on finding any installation,
setup, or environment configuration steps, even informal ones (e.g. listing required
Python version, package names without explicit commands).

Respond with a JSON object ONLY:
{
  "has_installation_instructions": true | false,
  "confidence": "high",
  "evidence": "<one decisive sentence>",
  "instruction_types": ["list", "of", "found", "types"],
  "installation_file": "<file path or null>"
}
Return ONLY valid JSON.
"""

# Targeted prompt for vague READMEs that mention prereqs but have no commands
IMPLICIT_INSTALL_PROMPT = """\
You are an expert code reviewer. Apply a BROAD definition of installation instructions:
- Listing required packages or Python/library versions by name = YES (type: "other")
- Mentioning "you need X installed" = YES (type: "other")
- A requirements.txt or environment.yml being present = YES (type: "requirements file")
- Only a paper abstract or citation info with zero technical content = NO

Be decisive. Return "high" confidence unless genuinely mid-truncation.

Respond with a JSON object ONLY:
{
  "has_installation_instructions": true | false,
  "confidence": "high",
  "evidence": "<one decisive sentence>",
  "instruction_types": ["list", "of", "found", "types"],
  "installation_file": "<file path or null>"
}
Return ONLY valid JSON.
"""


def diagnose_result(result, repo_content):
    """Return a list of matching diagnosis IDs for a medium/low confidence result."""
    return diagnose_with_rules(result, repo_content, DIAGNOSIS_RULES)


def heal_result(result, repo_content, owner, repo):
    """
    Attempt to fix a medium/low confidence result.
    Returns an updated result dict (or the original if healing failed).
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


def print_confidence_report(results, repo_contents):
    """Print a confidence breakdown with root-cause diagnosis for non-high results."""
    print_shared_confidence_report(results, repo_contents, diagnose_result)



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



def save_results(results, path=None):
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

    status_colors = {
        "yes":     "C6EFCE",  # green
        "no":      "FFDDC1",  # orange
        "skipped": "FFEB9C",  # yellow
        "error":   "FFC7CE",  # red
    }

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
        fill_color = status_colors.get(r.get("status"), "FFFFFF")
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
    yes     = sum(1 for r in results if r.get("status") == "yes")
    no      = sum(1 for r in results if r.get("status") == "no")
    skipped = sum(1 for r in results if r.get("status") == "skipped")
    errors  = sum(1 for r in results if r.get("status") == "error")
    total   = len(results)

    summary_data = [
        ("Total Repos Checked", total),
        ("Have Installation Instructions", yes),
        ("Missing Installation Instructions", no),
        ("Skipped (non-GitHub)", skipped),
        ("Errors", errors),
        ("Coverage (%)", f"={yes}/{total if total else 1}*100"),
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
    print_confidence_report(results, repo_contents)

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