"""Rule-based API documentation checker that uses the GitHub API only.

This is a deterministic companion to the LLM-based API-documentation checker
in this folder.  No LLM, no token cost, fully reproducible.

It answers the same core question for one repository:

  Does the repo contain API documentation, either as generated docs/configs
  or as source-level docstrings / structured API comments?

How it works (pipeline):
  1. ``list_all_repo_files`` — one GitHub Trees API call gets the whole
     file listing.
  2. ``_collect_candidate_paths`` — buckets the listing into doc-config
     files, doc pages, READMEs, and source files (with diversified sampling
     so one folder doesn't dominate).
  3. For each bucket, fetch + analyse:
       a) doc-config buckets via ``_doc_tool_from_named_file`` /
          ``_doc_tool_from_content`` (Sphinx, MkDocs, Doxygen, OpenAPI…).
       b) doc pages + READMEs via ``_analyse_doc_page`` (looks for API/
          reference markers).
       c) source files via ``_analyse_source_file`` (Python AST docstrings;
          regex-based detectors for Java/Go/Rust/C#/JS/TS/C++).
  4. Synthesise the verdict, evidence text, and per-language coverage.

Heuristics used:
  1. Generated docs / API-spec files:
       Sphinx (incl. MyST markdown), MkDocs, Doxygen, JSDoc, TypeDoc,
       ReadTheDocs, pdoc, OpenAPI / Swagger.
  2. API/reference documentation pages:
       docs/api/, docs/reference/, api_docs/, javadoc/, and similar folders,
       confirmed by API/reference markers in the file content.
  3. Source-level documentation:
       Python docstrings via ``ast.get_docstring``.
       Structured API comments for other languages via Javadoc / Doxygen /
       XML-doc / Rust / Go style regexes.

CLI modes:
  python api_documentation_check_no_llm_used.py            → batch over PAPERS
  python api_documentation_check_no_llm_used.py owner/repo → single repo, JSON

The output mirrors the LLM checker contract as closely as possible, plus
two extra fields that only the rule-based checker can compute cheaply:

  {
    "has_api_documentation": true | false | null,
    "evidence": "...",
    "documentation_types": [...],
    "doc_tool": "Sphinx | MkDocs | ... | none",
    "doc_files": [
      {"path": "...", "type": "...", "description": "..."}
    ],

    # Python public-API coverage (None if the repo has no Python sources).
    # Skips private (_foo) and name-mangled (__foo) names; counts dunders.
    "public_api_summary": {
      "total_public": int,
      "documented_public": int,
      "high_quality_documented": int,   # subset with Args/Returns/Raises sections
      "coverage_percent": float | null
    } | null,

    # If GitHub Pages is configured for this repo, the published URL.
    "pages_url": "https://owner.github.io/repo" | null
  }
"""

import argparse
import ast
import json
import re
import sys
import time
from pathlib import Path

from openpyxl import Workbook

# Add parent directories to Python path for local imports.
# Folder layout for path resolution:
#   _this.parent                       = check_github_repo_api_documentation/
#   _this.parent.parent                = 5.1.4.code_documentation_quality/
#   _this.parent.parent.parent         = 5.1.code_availability/
#   _this.parent.parent.parent.parent  = requirement_checks/
_this = Path(__file__).resolve()
sys.path.insert(0, str(_this.parent.parent.parent.parent))

from common.github_helpers import (
    load_dotenv, list_all_repo_files, fetch_file_content, parse_github_repo,
    is_github, github_get,
)
from common.excel_output import (
    count_statuses,
    write_header_row, write_results_data_rows, write_summary_sheet,
    thin_border, auto_row_height,
)
from config import (
    SOURCE_CODE_EXTENSIONS,
    SKIP_FOLDER_PREFIXES,
    SKIP_BASENAMES,
    TARGET_DOC_FILENAMES,
    DOC_FOLDER_PREFIXES,
    DOC_PAGE_EXTENSIONS,
    MAX_SOURCE_FILES,
    MAX_DOC_FILES,
)
from data.papers_from_database import PAPERS

load_dotenv()

DOC_PAGE_EXTENSIONS_WITH_HTML = DOC_PAGE_EXTENSIONS | {".html", ".htm"}

DOC_TOOL_PRIORITY = [
    "Sphinx",
    "MkDocs",
    "Doxygen",
    "Javadoc",
    "JSDoc",
    "TypeDoc",
    "pdoc",
    "OpenAPI",
    "other",
]

DOCUMENTATION_TYPE_PRIORITY = [
    "python docstrings",
    "java javadoc",
    "c++ doxygen comments",
    "c# xml docs",
    "go doc comments",
    "rust doc comments",
    "sphinx docs",
    "mkdocs docs",
    "doxygen docs",
    "jsdoc/typedoc",
    "readthedocs integration",
    "openapi/swagger spec",
    "pdoc/pydoc",
    "other",
]

PYTHON_PARAM_MARKERS_RE = re.compile(
    r"\b(args?|parameters?|params?|returns?|yields?|raises?|attributes?)\b",
    re.IGNORECASE,
)
RST_AUTODOC_RE = re.compile(
    r"(?im)^\.\.\s+(automodule|autoclass|autofunction|autodata|autoproperty)::"
)
DOC_PAGE_MARKERS_RE = re.compile(
    r"(?i)(api reference|reference guide|module reference|class reference|"
    r"function reference|method reference|endpoint reference|endpoints|"
    r"parameters?|returns?|attributes?|swagger ui|openapi)"
)
DOC_PAGE_HEADING_RE = re.compile(
    r"(?im)^\s{0,3}(?:#|=|-)+\s*(api|reference|api reference|module reference|"
    r"class reference|function reference|endpoints?)\b"
)
HTML_DOC_MARKERS_RE = re.compile(
    r"(?i)(genindex|py-modindex|api reference|class index|module index|swagger-ui)"
)
STRUCTURED_BLOCK_RE = re.compile(
    r"/\*\*[\s\S]{0,600}?(?:@param|@return|@brief|@tparam|@throws|@type|"
    r"<summary>|<param|<returns>)",
    re.IGNORECASE,
)
PLAIN_BLOCK_WITH_DECL_RE = re.compile(
    r"/\*\*[\s\S]{5,300}?\*/\s*(?:export\s+)?(?:public|private|protected|"
    r"static|async|inline|virtual|final|class|interface|struct|enum|function|"
    r"func|type|impl|pub)\b",
    re.IGNORECASE,
)
CSHARP_XML_DOC_RE = re.compile(r"(?im)^\s*///\s*<(summary|param|returns|remarks)>")
RUST_DOC_RE = re.compile(r"(?im)^\s*(///|//! )")
GO_DOC_RE = re.compile(
    r"(?im)^\s*//\s+([A-Z][A-Za-z0-9_]*)\b.*\n\s*(func|type|const|var)\s+\1\b"
)
CPP_DOXYGEN_RE = re.compile(r"(?im)^\s*///[</!]?")

PYPROJECT_TOOL_PATTERNS = [
    (re.compile(r"\[tool\.pdoc\]", re.IGNORECASE), "pdoc", "pdoc/pydoc"),
    (re.compile(r"\[tool\.mkdocs\]", re.IGNORECASE), "MkDocs", "mkdocs docs"),
    (re.compile(r"\btypedoc\b", re.IGNORECASE), "TypeDoc", "jsdoc/typedoc"),
    (re.compile(r"\bjsdoc\b", re.IGNORECASE), "JSDoc", "jsdoc/typedoc"),
    # MyST first so it isn't shadowed by the generic Sphinx pattern below.
    (re.compile(r"\b(myst[_-]?parser|myst[_-]?nb)\b", re.IGNORECASE), "Sphinx", "sphinx docs (myst)"),
    (re.compile(r"\b(sphinx|autodoc|napoleon)\b", re.IGNORECASE), "Sphinx", "sphinx docs"),
]

# Markdown variant of Sphinx autodoc — MyST renders ``` {eval-rst} ``` /
# ``` {autodoc2} ``` / ``` {autosummary} ``` directive blocks.
MYST_AUTODOC_RE = re.compile(
    r"```\{(eval-rst|autodoc2|autosummary|automodule|autoclass|autofunction)\b"
)


def _empty_result(*, has=False, evidence=""):
    """Return an empty result dict matching the public output schema.

    Used by every "we have nothing to report" code path so the schema stays
    consistent across success, failure, and skipped cases.

    Args:
        has:       Value for ``has_api_documentation`` (False, None for
                   error/skipped, True is unused here).
        evidence:  Free-form string explaining why the result is empty.
    """
    return {
        "has_api_documentation": has,
        "evidence": evidence,
        "documentation_types": [],
        "doc_tool": "none",
        "doc_files": [],
        "public_api_summary": None,
        "pages_url": None,
    }


def _is_source_file(path_lower, basename):
    """Return True if the file is a source file worth scanning."""
    ext = "." + basename.rsplit(".", 1)[-1] if "." in basename else ""
    if ext not in SOURCE_CODE_EXTENSIONS:
        return False
    if any(path_lower.startswith(prefix) for prefix in SKIP_FOLDER_PREFIXES):
        return False
    if basename in SKIP_BASENAMES:
        return False
    return True


def _diversified_source_sample(paths, max_files):
    """Pick up to ``max_files`` source paths spread across top-level folders.

    Naive truncation (``paths[:max_files]``) often produces a sample dominated
    by one large folder (e.g. ``src/`` or ``server/``) which biases the
    docstring detection.  Round-robin selection across top-level buckets gives
    every part of the repo a fair chance of being represented.
    """
    if not paths or max_files <= 0:
        return []

    buckets = {}
    for path in paths:
        top = "" if "/" not in path else path.split("/", 1)[0]
        buckets.setdefault(top, []).append(path)

    for key in buckets:
        buckets[key].sort(key=str.lower)

    ordered_keys = [""] if "" in buckets else []
    ordered_keys += sorted((key for key in buckets if key != ""), key=str.lower)

    selected = []
    round_idx = 0
    while len(selected) < max_files:
        added_this_round = False
        for key in ordered_keys:
            if round_idx < len(buckets[key]):
                selected.append(buckets[key][round_idx])
                added_this_round = True
                if len(selected) >= max_files:
                    break
        if not added_this_round:
            break
        round_idx += 1

    return selected


def _sort_documentation_types(doc_types):
    """Return documentation types in a stable, reviewer-friendly order."""
    seen = set(doc_types)
    ordered = [value for value in DOCUMENTATION_TYPE_PRIORITY if value in seen]
    extras = sorted(seen - set(ordered))
    return ordered + extras


def _pick_primary_tool(tools_found):
    """Pick the highest-priority tool from ``DOC_TOOL_PRIORITY``.

    A repo can show evidence of more than one tool (e.g. Sphinx config plus
    a stray Doxygen file).  We surface a single "primary" tool in the output
    by walking ``DOC_TOOL_PRIORITY`` top-down and returning the first match.
    Returns ``"none"`` when no tools were discovered.
    """
    if not tools_found:
        return "none"
    for candidate in DOC_TOOL_PRIORITY:
        if candidate in tools_found:
            return candidate
    return next(iter(tools_found))


def _add_doc_file(entry_map, path, entry_type, description):
    """Add one documentation evidence entry, deduplicated by path and type."""
    key = (path, entry_type)
    if key not in entry_map:
        entry_map[key] = {
            "path": path,
            "type": entry_type,
            "description": description,
        }


def _doc_tool_from_named_file(path_lower):
    """Classify a known documentation/config file by filename."""
    basename = path_lower.split("/")[-1]

    if basename in {"mkdocs.yml", "mkdocs.yaml"}:
        return "MkDocs", "mkdocs docs", "generated_docs_config", "MkDocs configuration file for generated documentation."
    if basename in {"doxyfile", "doxyfile.in"}:
        return "Doxygen", "doxygen docs", "generated_docs_config", "Doxygen configuration file for generated API documentation."
    if basename in {"jsdoc.json", ".jsdoc.json"}:
        return "JSDoc", "jsdoc/typedoc", "generated_docs_config", "JSDoc configuration file for generated API documentation."
    if basename == "typedoc.json":
        return "TypeDoc", "jsdoc/typedoc", "generated_docs_config", "TypeDoc configuration file for generated API documentation."
    if basename in {"openapi.yaml", "openapi.yml", "openapi.json", "swagger.yaml", "swagger.yml", "swagger.json", "api.yaml", "api.yml", "api.json"}:
        return "OpenAPI", "openapi/swagger spec", "api_spec", "OpenAPI / Swagger specification file describing the API surface."
    if basename in {".readthedocs.yml", ".readthedocs.yaml", "readthedocs.yml", "readthedocs.yaml"}:
        return "other", "readthedocs integration", "generated_docs_config", "ReadTheDocs configuration file for published documentation."
    if path_lower in {"conf.py", "docs/conf.py", "index.rst", "docs/index.rst", "docs/index.md"}:
        return "Sphinx", "sphinx docs", "generated_docs_config", "Sphinx documentation entry/configuration file."

    return None, None, None, None


def _doc_tool_from_content(path_lower, content):
    """Detect documentation tooling from file content when the name alone is ambiguous."""
    if not content:
        return None, None, None, None

    if path_lower.endswith("pyproject.toml"):
        for pattern, tool, doc_type in PYPROJECT_TOOL_PATTERNS:
            if pattern.search(content):
                return tool, doc_type, "generated_docs_config", f"pyproject.toml configures {tool} documentation tooling."

    if path_lower.endswith("conf.py") or path_lower.endswith("index.rst"):
        if RST_AUTODOC_RE.search(content) or "sphinx" in content.lower():
            doc_type = "sphinx docs"
            description = "Sphinx configuration/content with autodoc-style API directives."
            if "myst_parser" in content or "myst_nb" in content:
                doc_type = "sphinx docs (myst)"
                description = "Sphinx configuration enabling MyST markdown parsing."
            return "Sphinx", doc_type, "generated_docs_config", description

    return None, None, None, None


def _analyse_doc_page(path, content):
    """Return doc-page evidence when the file looks like API/reference docs.

    A file qualifies as an API/reference doc page if either:
      - it lives in a strong path (``docs/api/``, ``docs/reference/``,
        ``apidoc/``, etc.) — folder name alone is enough; OR
      - its content scores at least 2 marker hits (API/reference vocabulary,
        autodoc directives, HTML index markers).  Two markers is the
        threshold because any single signal is too easy to trigger
        accidentally (e.g. the word "parameters" appears in tutorials).

    Returns a dict shaped like other ``_add_doc_file`` entries, or None when
    the file isn't a clear API doc page.
    """
    if not content:
        return None

    path_lower = path.lower()
    strong_path = path_lower.startswith((
        "docs/api/",
        "docs/reference/",
        "docs/api-reference/",
        "apidoc/",
        "api_docs/",
        "api-docs/",
        "javadoc/",
    ))

    basename = path_lower.split("/")[-1]
    marker_hits = 0
    if DOC_PAGE_MARKERS_RE.search(content):
        marker_hits += 1
    if DOC_PAGE_HEADING_RE.search(content):
        marker_hits += 1
    if RST_AUTODOC_RE.search(content) or MYST_AUTODOC_RE.search(content):
        marker_hits += 2
    if basename.endswith((".html", ".htm")) and HTML_DOC_MARKERS_RE.search(content):
        marker_hits += 1

    if not strong_path and marker_hits < 2:
        return None

    description = "Documentation page with API/reference structure or autodoc markers."
    if strong_path:
        description = "Documentation page inside an API/reference documentation folder."

    return {
        "type": "doc_page",
        "description": description,
        "documentation_type": "other",
    }


def _is_public_python_name(name):
    """Public-API filter for Python identifiers used in coverage stats.

    Conventions matched:
      - ``foo``       → public.
      - ``__init__``  → public (constructor, conventionally documented).
      - ``__call__`` etc → public (dunders are part of the public surface).
      - ``_helper``   → private, skipped.
      - ``__mangled`` (no trailing __) → name-mangled private, skipped.
    """
    if not name:
        return False
    if name.startswith("__") and name.endswith("__"):
        return True
    return not name.startswith("_")


def _count_python_docstrings(content):
    """Return docstring + public-API counts for Python source.

    The result includes both *documented* and *total* counts so callers can
    compute coverage percentages.  Only public functions/classes (see
    ``_is_public_python_name``) are counted toward coverage so that helper
    methods and private internals don't dilute the figure.

    Returns None on syntax errors.
    """
    try:
        tree = ast.parse(content)
    except SyntaxError:
        return None

    counts = {
        # Documented (has any docstring)
        "module": 1 if ast.get_docstring(tree) else 0,
        "class": 0,
        "function": 0,
        "param_like": 0,         # subset of documented with Args/Returns/Raises sections
        # Totals over the public surface (used for coverage %)
        "public_class_total": 0,
        "public_function_total": 0,
    }

    if counts["module"]:
        module_doc = ast.get_docstring(tree) or ""
        if PYTHON_PARAM_MARKERS_RE.search(module_doc):
            counts["param_like"] += 1

    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef):
            if not _is_public_python_name(node.name):
                continue
            counts["public_class_total"] += 1
            doc = ast.get_docstring(node)
            if doc:
                counts["class"] += 1
                if PYTHON_PARAM_MARKERS_RE.search(doc):
                    counts["param_like"] += 1
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            if not _is_public_python_name(node.name):
                continue
            counts["public_function_total"] += 1
            doc = ast.get_docstring(node)
            if doc:
                counts["function"] += 1
                if PYTHON_PARAM_MARKERS_RE.search(doc):
                    counts["param_like"] += 1

    return counts


def _analyse_python_source(path, content):
    """Inspect one Python file for meaningful docstrings.

    Returns None when the file has no documented public API at all.
    Otherwise returns a dict with:

      * ``path``, ``entry_type`` (``docstring_function`` / ``_class`` /
        ``_module``), ``description``, ``documentation_type``.
      * ``strength`` — ``"strong"`` if any function/class is documented or
        any docstring uses Args/Returns/Raises sections; ``"weak"`` if only
        a module-level docstring without param markers is present.  The
        orchestrator uses the strength to decide whether one Python file
        alone is enough evidence of API docs.
      * ``public_total`` / ``public_documented`` / ``high_quality_documented``
        — coverage stats over the public surface for the orchestrator to
        sum across files.
    """
    counts = _count_python_docstrings(content)
    if not counts:
        return None

    total_structured = counts["function"] + counts["class"]
    if total_structured == 0 and counts["module"] == 0:
        return None

    if counts["function"] > 0:
        entry_type = "docstring_function"
    elif counts["class"] > 0:
        entry_type = "docstring_class"
    else:
        entry_type = "docstring_module"

    description_bits = []
    if counts["function"]:
        description_bits.append(f"{counts['function']} function docstrings")
    if counts["class"]:
        description_bits.append(f"{counts['class']} class docstrings")
    if counts["module"]:
        description_bits.append("module docstring")

    strength = "strong" if total_structured > 0 or counts["param_like"] > 0 else "weak"
    description = "Contains " + ", ".join(description_bits) + "."

    return {
        "path": path,
        "entry_type": entry_type,
        "description": description,
        "documentation_type": "python docstrings",
        "strength": strength,
        # Coverage stats for the orchestrator to aggregate.
        "public_total": counts["public_function_total"] + counts["public_class_total"],
        "public_documented": counts["function"] + counts["class"],
        "high_quality_documented": counts["param_like"],
    }


def _analyse_non_python_source(path, content):
    """Inspect one non-Python source file for structured API comments.

    Per-language detectors:
      * ``.cs``                     → C# XML doc tags (``/// <summary>`` …)
      * ``.go``                     → Doc comment on exported identifier
      * ``.rs``                     → Rust ``///`` and ``//!`` lines (≥2)
      * ``.java`` / ``.scala`` / ``.kt`` → Javadoc ``/** */`` blocks
      * ``.js`` / ``.ts`` / ``.jsx`` / ``.tsx`` → JSDoc / TypeDoc blocks
      * ``.cpp`` / ``.c`` / ``.h`` / ``.hpp`` / ``.cc`` / ``.cxx`` /
        ``.cu`` / ``.cuh``           → Doxygen blocks + ``///`` line docs

    Returns the same shape as ``_analyse_python_source`` (path, entry_type,
    description, documentation_type, strength) — but without the coverage
    fields, since reliable "total public API" counts require proper parsers
    we don't have here.  Returns None when nothing matches.
    """
    path_lower = path.lower()
    ext = "." + path_lower.rsplit(".", 1)[-1] if "." in path_lower else ""

    if ext == ".cs":
        count = len(CSHARP_XML_DOC_RE.findall(content))
        if count:
            return {
                "path": path,
                "entry_type": "other",
                "description": f"Contains {count} XML documentation comment blocks.",
                "documentation_type": "c# xml docs",
                "strength": "strong",
            }

    if ext == ".go":
        count = len(GO_DOC_RE.findall(content))
        if count:
            return {
                "path": path,
                "entry_type": "other",
                "description": f"Contains {count} exported-identifier doc comment patterns.",
                "documentation_type": "go doc comments",
                "strength": "strong",
            }

    if ext == ".rs":
        count = len(RUST_DOC_RE.findall(content))
        if count >= 2:
            return {
                "path": path,
                "entry_type": "other",
                "description": f"Contains {count} Rust doc comment lines.",
                "documentation_type": "rust doc comments",
                "strength": "strong",
            }

    structured_blocks = len(STRUCTURED_BLOCK_RE.findall(content))
    plain_blocks = len(PLAIN_BLOCK_WITH_DECL_RE.findall(content))
    cpp_line_docs = len(CPP_DOXYGEN_RE.findall(content)) if ext in {".cpp", ".c", ".h", ".hpp", ".cc", ".cxx", ".cu", ".cuh"} else 0

    if ext in {".java", ".scala", ".kt"} and (structured_blocks or plain_blocks):
        count = structured_blocks or plain_blocks
        return {
            "path": path,
            "entry_type": "other",
            "description": f"Contains {count} Javadoc-style documentation blocks.",
            "documentation_type": "java javadoc",
            "strength": "strong",
        }

    if ext in {".js", ".ts", ".jsx", ".tsx"} and (structured_blocks or plain_blocks):
        count = structured_blocks or plain_blocks
        return {
            "path": path,
            "entry_type": "other",
            "description": f"Contains {count} JSDoc / TypeDoc documentation blocks.",
            "documentation_type": "jsdoc/typedoc",
            "strength": "strong",
        }

    if ext in {".cpp", ".c", ".h", ".hpp", ".cc", ".cxx", ".cu", ".cuh"} and (structured_blocks or cpp_line_docs):
        count = structured_blocks + cpp_line_docs
        return {
            "path": path,
            "entry_type": "other",
            "description": f"Contains {count} Doxygen-style API comment markers.",
            "documentation_type": "c++ doxygen comments",
            "strength": "strong" if structured_blocks else "weak",
        }

    return None


def _analyse_source_file(path, content):
    """Dispatch source-file analysis by language."""
    if path.lower().endswith(".py"):
        return _analyse_python_source(path, content)
    return _analyse_non_python_source(path, content)


def _summarize_python_coverage(source_hits):
    """Aggregate Python public-API coverage across analysed source files.

    Returns None when none of the analysed files were Python (so the field is
    omitted gracefully for non-Python repos).  Otherwise returns a dict with
    the totals, documented counts, high-quality counts, and the resulting
    coverage percent.
    """
    py_hits = [h for h in source_hits if h.get("documentation_type") == "python docstrings"]
    if not py_hits:
        return None

    total = sum(h.get("public_total", 0) for h in py_hits)
    documented = sum(h.get("public_documented", 0) for h in py_hits)
    high_quality = sum(h.get("high_quality_documented", 0) for h in py_hits)

    coverage_percent = round((documented / total) * 100, 1) if total else None
    return {
        "total_public": total,
        "documented_public": documented,
        "high_quality_documented": high_quality,
        "coverage_percent": coverage_percent,
    }


def _check_github_pages(owner, repo):
    """Return the published GitHub Pages URL for the repo, or None.

    Calls ``GET /repos/{owner}/{repo}/pages``.  GitHub returns 404 when the
    repo has no Pages site (which ``github_get`` surfaces as None), so a
    truthy return tells us docs are actually deployed and reachable.
    """
    data = github_get(f"repos/{owner}/{repo}/pages")
    if not data:
        return None
    return data.get("html_url") or None


def _collect_candidate_paths(all_files):
    """Split the repo file listing into four prioritised buckets.

    Each file goes into exactly one bucket based on a priority order:

      1. ``readme_paths`` — anything whose basename starts with ``readme``.
      2. ``doc_config_paths`` — known doc/spec config files
         (mkdocs.yml, conf.py, openapi.yaml, pyproject.toml, …).
      3. ``doc_page_paths`` — files inside ``DOC_FOLDER_PREFIXES``
         (``docs/``, ``docs/api/``, ``javadoc/`` …) whose extension marks
         them as a doc page (.md, .rst, .txt, .html).
      4. ``source_paths`` — anything that ``_is_source_file`` accepts.

    Returns the four buckets, with the last three pre-trimmed to caps
    (``MAX_DOC_FILES``, ``MAX_SOURCE_FILES``) and the source bucket
    diversified across top-level folders.
    """
    doc_config_paths = []
    doc_page_paths = []
    readme_paths = []
    source_paths = []

    for file_info in all_files:
        path = file_info["path"]
        path_lower = path.lower()
        basename = path_lower.split("/")[-1]
        ext = "." + basename.rsplit(".", 1)[-1] if "." in basename else ""

        if basename.startswith("readme"):
            readme_paths.append(path)
        elif path_lower in TARGET_DOC_FILENAMES or basename in TARGET_DOC_FILENAMES:
            doc_config_paths.append(path)
        elif any(path_lower.startswith(prefix) for prefix in DOC_FOLDER_PREFIXES):
            if ext in DOC_PAGE_EXTENSIONS_WITH_HTML:
                doc_page_paths.append(path)
        elif _is_source_file(path_lower, basename):
            source_paths.append(path)

    doc_config_paths.sort(key=lambda value: (value.count("/"), value.lower()))
    doc_page_paths.sort(key=lambda value: (value.count("/"), value.lower()))
    readme_paths.sort(key=str.lower)

    return doc_config_paths, doc_page_paths[:MAX_DOC_FILES], readme_paths[:1], _diversified_source_sample(source_paths, MAX_SOURCE_FILES)


def check_api_documentation_without_llm(owner, repo):
    """Check one GitHub repo for API documentation. The main public entry point.

    Pipeline:
      1. List every file via the GitHub Trees API.
      2. Bucket the listing (config / doc page / README / source).
      3. Phase A — for each doc-config path, fetch only when the filename
         alone isn't decisive (e.g. ``pyproject.toml``, ``conf.py``).
      4. Phase B — for each doc page + first README, fetch and apply the
         API/reference marker check.
      5. Phase C — for each source path, fetch and run the language-specific
         docstring/comment detector.
      6. Synthesis: a repo "has API documentation" if EITHER phase A/B found
         a generated-docs signal OR phase C found at least one strong source
         file (or two weak ones).
      7. Append Python coverage stats and the GitHub Pages URL (if any).

    Returns a dict matching the schema in the module docstring.  All fields
    are always present so callers can assume the contract.
    """
    all_files = list_all_repo_files(owner, repo)
    if not all_files:
        return _empty_result(
            has=False,
            evidence="No repository files were returned by the GitHub API.",
        )

    doc_config_paths, doc_page_paths, readme_paths, source_paths = _collect_candidate_paths(all_files)

    entry_map = {}
    documentation_types = set()
    tools_found = []
    strong_source_hits = []
    weak_source_hits = []

    for path in doc_config_paths:
        path_lower = path.lower()
        content = None
        tool, doc_type, entry_type, description = _doc_tool_from_named_file(path_lower)
        if tool is None or path_lower.endswith(("pyproject.toml", "conf.py", "index.rst")):
            content = fetch_file_content(owner, repo, path)
            tool2, doc_type2, entry_type2, description2 = _doc_tool_from_content(path_lower, content or "")
            if tool2 is not None:
                tool, doc_type, entry_type, description = tool2, doc_type2, entry_type2, description2

        if tool is None:
            continue

        if tool not in tools_found:
            tools_found.append(tool)
        documentation_types.add(doc_type)
        _add_doc_file(entry_map, path, entry_type, description)

    for path in doc_page_paths + readme_paths:
        content = fetch_file_content(owner, repo, path)
        page_result = _analyse_doc_page(path, content or "")
        if page_result is None:
            continue
        documentation_types.add(page_result["documentation_type"])
        _add_doc_file(entry_map, path, page_result["type"], page_result["description"])

    for path in source_paths:
        content = fetch_file_content(owner, repo, path)
        if not content:
            continue
        source_result = _analyse_source_file(path, content)
        if source_result is None:
            continue
        if source_result["strength"] == "strong":
            strong_source_hits.append(source_result)
        else:
            weak_source_hits.append(source_result)

    has_source_docstrings = bool(strong_source_hits) or len(strong_source_hits) + len(weak_source_hits) >= 2
    has_generated_docs = any(
        entry["type"] in {"generated_docs_config", "api_spec", "doc_page"}
        for entry in entry_map.values()
    )

    if has_source_docstrings:
        for hit in strong_source_hits + weak_source_hits:
            documentation_types.add(hit["documentation_type"])
            _add_doc_file(entry_map, hit["path"], hit["entry_type"], hit["description"])

    has_api_documentation = has_generated_docs or has_source_docstrings
    doc_files = sorted(entry_map.values(), key=lambda entry: (entry["path"].lower(), entry["type"]))
    doc_tool = _pick_primary_tool(tools_found)

    if has_api_documentation:
        evidence_parts = []
        if has_generated_docs:
            generated_paths = [entry["path"] for entry in doc_files if entry["type"] in {"generated_docs_config", "api_spec", "doc_page"}]
            preview = ", ".join(generated_paths[:3])
            if doc_tool != "none":
                evidence_parts.append(f"generated documentation signals via {doc_tool} ({preview})")
            else:
                evidence_parts.append(f"generated documentation pages/specs ({preview})")
        if has_source_docstrings:
            source_hits = strong_source_hits + weak_source_hits
            preview = ", ".join(hit["path"] for hit in source_hits[:3])
            evidence_parts.append(f"structured source docstrings/comments in {len(source_hits)} files ({preview})")
        evidence = "Found " + " and ".join(evidence_parts) + "."
    else:
        evidence = "No generated documentation config, API reference pages, API specs, or sufficiently strong source docstrings were detected."
        doc_files = []
        documentation_types = set()
        doc_tool = "none"

    public_api_summary = _summarize_python_coverage(strong_source_hits + weak_source_hits)
    pages_url = _check_github_pages(owner, repo)

    return {
        "has_api_documentation": has_api_documentation,
        "evidence": evidence,
        "documentation_types": _sort_documentation_types(documentation_types),
        "doc_tool": doc_tool,
        "doc_files": doc_files,
        "public_api_summary": public_api_summary,
        "pages_url": pages_url,
    }


def check_api_documentation_from_url(repo_url):
    """Parse a GitHub URL and run the rule-based checker."""
    owner, repo = parse_github_repo(repo_url)
    if not owner:
        return _empty_result(
            has=None,
            evidence=f"Could not parse GitHub repository URL: {repo_url}",
        )
    return check_api_documentation_without_llm(owner, repo)


# =============================================================================
# BATCH MODE — run over all PAPERS and save Excel
# =============================================================================

def _check_paper(paper):
    """Run the no-LLM checker on one paper entry from papers_from_database.

    Returns a result dict shaped for the Excel writer:
      title, repo, status (yes/no/skipped/error), note, plus all fields
      from check_api_documentation_without_llm.

    Mapping rules:
      - URL is not GitHub                       → status="skipped"
      - URL parses but API returned no files    → status="error"
      - has_api_documentation True / False      → status="yes" / "no"
    """
    url   = paper.get("repo", "") or ""
    title = paper.get("title", "")

    if not is_github(url):
        return {
            "title": title, "repo": url,
            "status": "skipped",
            "note": f"Not a GitHub repo: {url}" if url else "No repo URL provided",
            **_empty_result(has=None),
        }

    try:
        result = check_api_documentation_from_url(url)
    except Exception as exc:
        return {
            "title": title, "repo": url,
            "status": "error",
            "note": f"Unexpected error: {exc}",
            **_empty_result(has=None),
        }

    has = result.get("has_api_documentation")
    if has is None:
        status, note = "error", result.get("evidence", "")
    elif has:
        status, note = "yes", ""
    else:
        status, note = "no", ""

    return {
        "title": title,
        "repo": url,
        "status": status,
        "note": note,
        **result,
    }


def _save_batch_to_excel(results, output_path):
    """Write batch results to an Excel workbook with Results + Summary sheets."""
    wb     = Workbook()
    border = thin_border()

    # ── Results sheet ────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Results"

    headers = [
        "#", "Status", "Title", "Repo",
        "Has API Docs", "Doc Tool", "Documentation Types",
        "Coverage %", "Pages URL", "Doc Files", "Evidence", "Note",
    ]
    widths = [5, 10, 45, 40, 14, 14, 32, 12, 35, 50, 55, 35]
    write_header_row(ws, headers, widths, fill_hex="2F5496", border=border)

    def row_data(r, num):
        summary = r.get("public_api_summary") or {}
        coverage = summary.get("coverage_percent")
        coverage_str = f"{coverage}%" if coverage is not None else ""
        return [
            num,
            (r.get("status") or "").upper(),
            r.get("title") or "",
            r.get("repo") or "",
            "Yes" if r.get("has_api_documentation") else
            ("No" if r.get("has_api_documentation") is False else "N/A"),
            r.get("doc_tool") or "",
            ", ".join(r.get("documentation_types") or []),
            coverage_str,
            r.get("pages_url") or "",
            "\n".join(f"{f.get('path')} ({f.get('type')})" for f in r.get("doc_files") or []),
            r.get("evidence") or "",
            r.get("note") or "",
        ]

    write_results_data_rows(
        ws, results, row_data,
        border=border,
        row_height_fn=lambda vals: auto_row_height(vals),
    )

    # ── Summary sheet ────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    write_summary_sheet(
        ws2, results,
        positive_label="Has API Documentation",
        negative_label="No API Documentation",
        fill_hex="2F5496",
        border=border,
    )

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"Saved to {output_path}")


def run_batch():
    """Run the no-LLM checker on every paper in PAPERS and save Excel."""
    print(f"Checking {len(PAPERS)} papers for API documentation (no LLM)...\n")
    results = []
    started = time.time()

    for i, paper in enumerate(PAPERS, 1):
        owner, repo = parse_github_repo(paper.get("repo", "") or "")
        label = f"{owner}/{repo}" if owner else (paper.get("repo") or "?")
        print(f"[{i:>2}/{len(PAPERS)}] {label} ...", end=" ", flush=True)

        result = _check_paper(paper)
        results.append(result)
        print(result["status"].upper())
        time.sleep(0.3)

    elapsed = time.time() - started
    counts = count_statuses(results)
    print(
        f"\nDone in {elapsed:.1f}s — "
        f"yes: {counts.get('yes', 0)}, no: {counts.get('no', 0)}, "
        f"skipped: {counts.get('skipped', 0)}, errors: {counts.get('error', 0)}"
    )

    output_path = _this.parent / "results" / "api_documentation_no_llm.xlsx"
    _save_batch_to_excel(results, output_path)


# =============================================================================
# CLI
# =============================================================================

def _parse_cli_repo(value):
    """Accept either a full GitHub URL or owner/repo on the command line."""
    if "github.com" in value:
        owner, repo = parse_github_repo(value)
    elif "/" in value:
        owner, repo = value.split("/", 1)
    else:
        owner, repo = None, None
    return owner, repo


def main():
    """CLI entry point.

    Two modes:
      * No argument → batch mode: runs over every paper in
        ``papers_from_database.PAPERS`` and writes results to an Excel file.
      * One argument → single-repo mode: checks one repo and prints JSON.
        The argument can be a full GitHub URL or ``owner/repo``.
    """
    parser = argparse.ArgumentParser(
        description="Check whether a GitHub repository contains API documentation without using an LLM."
    )
    parser.add_argument(
        "repo",
        nargs="?",
        default=None,
        help=(
            "GitHub repo URL or owner/repo (e.g. https://github.com/org/repo). "
            "If omitted, runs batch mode over all papers from papers_from_database.py."
        ),
    )
    args = parser.parse_args()

    if args.repo is None:
        run_batch()
        return

    owner, repo = _parse_cli_repo(args.repo)
    if not owner:
        raise SystemExit(f"Could not parse repo argument: {args.repo}")

    result = check_api_documentation_without_llm(owner, repo)
    print(json.dumps(result, indent=2, ensure_ascii=True))


if __name__ == "__main__":
    main()