"""Adapter layer between the Streamlit UI and the existing checker scripts.

Each checker is a standalone script that reads its papers from
``data.papers_source.load_papers()`` (which honours the ``PAPERS_JSON`` env var)
and writes an Excel file to a fixed location.  Rather than import the checkers
in-process (they reuse module names like ``config``/``prompts`` across folders,
which would collide), we run each as a **subprocess** with the upload injected
via an environment variable, then read back the Excel it produced.

This module is import-safe (no Streamlit) so it can be unit-tested on its own.
"""

import json
import os
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

REPO_ROOT = Path(__file__).resolve().parent.parent
RC = REPO_ROOT / "requirement_checks"
_Q1 = RC / "5.1.code_availability"
_Q2 = RC / "5.2.practitioner_usability_and_popularity"
_DOC = _Q1 / "5.1.4.code_documentation_quality"
_BEALLS = RC / "bealls_list_check"


@dataclass(frozen=True)
class Checker:
    id: str
    section: str            # "5.1" | "5.2" | "bealls"
    label: str
    script: Path            # script to run as a subprocess
    output: Path            # Excel it writes (we also scan its folder for fresher files)
    needs_llm: bool = False
    corpus_based: bool = False   # Beall's: input is Semantic Scholar records, fed via BEALLS_CORPUS_DIR
    note: str = ""


CHECKERS: List[Checker] = [
    # ── 5.1 Code availability & documentation ───────────────────────────────
    Checker("5.1.3.preprocessing", "5.1", "5.1.3 — Preprocessing / pipeline code (LLM)",
            _Q1 / "5.1.3.pre-processing_&_pipeline_code" / "check_paper_appendix_for_data_preprocessing_code.py",
            _Q1 / "5.1.3.pre-processing_&_pipeline_code" / "results" / "preprocessing_code_results.xlsx",
            needs_llm=True),
    Checker("5.1.4.inline", "5.1", "5.1.4 — Inline comments (LLM)",
            _DOC / "check_github_repo_inline_comments" / "check_github_repo_inline_comments.py",
            _DOC / "check_github_repo_inline_comments" / "results" / "inline_comments_results.xlsx",
            needs_llm=True),
    Checker("5.1.4.installation", "5.1", "5.1.4 — Installation instructions (LLM)",
            _DOC / "check_github_repo_installation_instructions" / "check_github_repo_installation_instructions.py",
            _DOC / "check_github_repo_installation_instructions" / "results" / "installation_instructions_results.xlsx",
            needs_llm=True),
    Checker("5.1.4.usage", "5.1", "5.1.4 — Usage / example commands (LLM)",
            _DOC / "check_github_repo_usage_examples" / "check_github_repo_example_commands.py",
            _DOC / "check_github_repo_usage_examples" / "results" / "usage_examples_results.xlsx",
            needs_llm=True),
    Checker("5.1.4.apidoc", "5.1", "5.1.4 — API documentation (LLM)",
            _DOC / "check_github_repo_api_documentation" / "check_github_repo_api_documentation.py",
            _DOC / "check_github_repo_api_documentation" / "results" / "api_documentation_results.xlsx",
            needs_llm=True),
    Checker("5.1.4.apidoc_norule", "5.1", "5.1.4 — API documentation (rule-based, no LLM)",
            _DOC / "check_github_repo_api_documentation" / "api_documentation_check_no_llm_used.py",
            _DOC / "check_github_repo_api_documentation" / "results" / "api_documentation_no_llm.xlsx"),
    Checker("5.1.5.license", "5.1", "5.1.5 — Code license (no LLM)",
            _Q1 / "5.1.5.code_license" / "5.1.5.code_license.py",
            _Q1 / "5.1.5.code_license" / "code_license.xlsx"),
    # ── 5.2 Practitioner usability & popularity ─────────────────────────────
    Checker("5.2.2.maintenance", "5.2", "5.2.2 — Maintenance activity indicators (no LLM)",
            _Q2 / "5.2.2.maintenance_activity_indicators" / "5.2.2.maintenance_activity_indicators.py",
            _Q2 / "5.2.2.maintenance_activity_indicators" / "maintenance_indicators.xlsx"),
    Checker("5.2.3.adoption", "5.2", "5.2.3 — Adoption metrics: stars/forks/PyPI (no LLM)",
            _Q2 / "5.2.3.adoption_metrics" / "5.2.3.adoption_metrics.py",
            _Q2 / "5.2.3.adoption_metrics" / "adoption_metrics.xlsx"),
    Checker("5.2.4.postpub", "5.2", "5.2.4 — Post-publication maintenance (no LLM)",
            _Q2 / "5.2.4.post_publication_maintenance" / "5.2.4.post_publication_maintenance.py",
            _Q2 / "5.2.4.post_publication_maintenance" / "post_publication_maintenance.xlsx"),
    # ── Beall's predatory-venue check (corpus = Semantic Scholar records) ────
    Checker("bealls.deterministic", "bealls", "Beall's check — deterministic (no LLM)",
            _BEALLS / "bealls_list_check.py",
            _BEALLS / "results" / "bealls_list_results.xlsx",
            corpus_based=True,
            note="Input is Semantic Scholar paper records (with publicationVenue), not GitHub links."),
    Checker("bealls.llm", "bealls", "Beall's check — with LLM second pass",
            _BEALLS / "bealls_llm_check.py",
            _BEALLS / "results" / "bealls_llm_results.xlsx",
            needs_llm=True, corpus_based=True,
            note="Input is Semantic Scholar paper records. Runs the deterministic pass, then the LLM pass."),
]

CHECKERS_BY_ID = {c.id: c for c in CHECKERS}


def checkers_for(section: str) -> List[Checker]:
    return [c for c in CHECKERS if c.section == section]


@dataclass
class RunResult:
    ok: bool
    returncode: int
    stdout: str
    stderr: str
    output_path: Optional[Path]
    output_bytes: Optional[bytes]


def _newest_xlsx_since(folder: Path, since_ts: float) -> Optional[Path]:
    """Newest .xlsx in *folder* modified at/after *since_ts* (handles the
    timestamped fallback Beall's writes when its default file is locked)."""
    if not folder.is_dir():
        return None
    candidates = [p for p in folder.glob("*.xlsx")
                  if not p.name.startswith("~$") and p.stat().st_mtime >= since_ts - 1]
    return max(candidates, key=lambda p: p.stat().st_mtime, default=None)


def run_checker(checker: Checker, items: list, *, aux_lists: Optional[dict] = None,
                extra_args: Optional[list] = None, env_overrides: Optional[dict] = None,
                timeout: Optional[int] = None) -> RunResult:
    """Run one checker over *items* and return its produced Excel.

    *items* is a list of paper objects (GitHub-link papers for 5.1/5.2; Semantic
    Scholar records for Beall's).  We inject them via a temp file + env var and
    run the checker as a subprocess, inheriting the repo's ``.env`` (the checker
    calls ``load_dotenv()`` itself, so GitHub/LLM keys are picked up).

    *aux_lists* maps a CLI flag to a list written to a temp JSON and passed to
    the checker — used for Beall's ``--whitelist`` / ``--blacklist``; empty
    lists are skipped.
    """
    started = time.time()
    with tempfile.TemporaryDirectory(prefix="bealls_ui_") as tmp:
        tmp = Path(tmp)
        env = os.environ.copy()
        env.update(env_overrides or {})
        if checker.corpus_based:
            # The corpus dir is globbed for *.json, so it must contain ONLY the
            # uploaded corpus — keep it in its own subdir away from the aux files.
            corpus_dir = tmp / "corpus"
            corpus_dir.mkdir()
            (corpus_dir / "uploaded_corpus.json").write_text(
                json.dumps(items, ensure_ascii=False), encoding="utf-8")
            env["BEALLS_CORPUS_DIR"] = str(corpus_dir)
        else:
            papers_file = tmp / "uploaded_papers.json"
            papers_file.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
            env["PAPERS_JSON"] = str(papers_file)

        cmd = [sys.executable, str(checker.script)]
        for flag, entries in (aux_lists or {}).items():
            if not entries:
                continue
            aux_file = tmp / (flag.lstrip("-") + ".json")   # in tmp, never the corpus dir
            aux_file.write_text(json.dumps(entries, ensure_ascii=False), encoding="utf-8")
            cmd += [flag, str(aux_file)]
        cmd += list(extra_args or [])

        proc = subprocess.run(cmd, env=env, cwd=str(REPO_ROOT),
                              capture_output=True, text=True, timeout=timeout)

    out_path = checker.output if checker.output.exists() else None
    fresh = _newest_xlsx_since(checker.output.parent, started)
    if fresh is not None:
        out_path = fresh

    output_bytes = out_path.read_bytes() if out_path else None
    return RunResult(
        ok=(proc.returncode == 0 and output_bytes is not None),
        returncode=proc.returncode,
        stdout=proc.stdout or "",
        stderr=proc.stderr or "",
        output_path=out_path,
        output_bytes=output_bytes,
    )


def normalize_upload(parsed) -> list:
    """Coerce uploaded/pasted JSON into a list of paper objects.

    Accepts a list of objects or a single object (wrapped). Raises ValueError
    with a friendly message otherwise.
    """
    if isinstance(parsed, dict):
        return [parsed]
    if isinstance(parsed, list):
        if not all(isinstance(x, dict) for x in parsed):
            raise ValueError("JSON list must contain paper objects (dicts).")
        return parsed
    raise ValueError("Upload must be a JSON list of papers or a single paper object.")
