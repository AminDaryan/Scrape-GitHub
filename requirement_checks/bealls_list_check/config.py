"""Configuration for the Beall's List predatory-venue check.

Paths and tunables live here so the matcher and the Excel writer stay focused
on logic.  The corpus folder lives outside the repo, so it is overridable via
the ``BEALLS_CORPUS_DIR`` environment variable.
"""

import os
from pathlib import Path

_HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Inputs / outputs
# ---------------------------------------------------------------------------

# Folder holding the Semantic Scholar dump JSON files (one list per file).
# Defaults to the in-repo docs folder; override with BEALLS_CORPUS_DIR if the
# data lives elsewhere.  _HERE.parent.parent is the repo root
# ("Paper repository scraping Code").
_DEFAULT_CORPUS_DIR = _HERE.parent.parent / "docs" / "Updated Abstract Papers"
CORPUS_DIR = Path(os.environ.get("BEALLS_CORPUS_DIR", str(_DEFAULT_CORPUS_DIR)))

# Glob for the corpus files inside CORPUS_DIR.
CORPUS_GLOB = "*.json"

# Vendored Beall's List snapshot produced by scrape_bealls_list.py.
SNAPSHOT_PATH = _HERE / "data" / "bealls_snapshot.json"

# Where the combined results workbook is written.
RESULTS_DIR = _HERE / "results"
RESULTS_FILENAME = "bealls_list_results.xlsx"

# ---------------------------------------------------------------------------
# Matching tunables
# ---------------------------------------------------------------------------

# difflib ratio threshold for the fuzzy "review" tier.  Deliberately high:
# in this corpus genuinely predatory venues are rare, so a loose threshold
# would surface far more false positives than true hits.
FUZZY_CUTOFF = 0.93

# Minimum normalized-name length to even attempt an exact-name match, so we
# never match on a stub like "ai" or "data".
MIN_NAME_LEN = 5

# Lists that the fuzzy tier is allowed to match against.  Fuzzy matching the
# low-relevance metric/vanity lists produces noise, so those are exact-only.
FUZZY_LIST_SOURCES = {"publishers", "standalone_journal", "hijacked"}

# Preprint-server hosts: a paper whose venue is one of these has no journal
# publisher and therefore cannot be "on Beall's List" — it is classified
# "no_venue" rather than clean/predatory.
PREPRINT_DOMAINS = {
    "arxiv.org", "biorxiv.org", "medrxiv.org", "chemrxiv.org",
    "preprints.org", "researchsquare.com", "ssrn.com", "techrxiv.org",
    "authorea.com", "osf.io", "hal.science", "hal.archives-ouvertes.fr",
    "eccc.weizmann.ac.il", "openreview.net",
}

# Normalized venue names that signal a preprint when no domain is available.
PREPRINT_NAMES = {
    "arxiv", "arxiv org", "biorxiv", "medrxiv", "chemrxiv", "ssrn",
    "research square", "researchsquare", "preprints", "preprints org",
    "techrxiv", "authorea", "openreview", "hal",
}

# ---------------------------------------------------------------------------
# Result statuses + Excel row colors
# ---------------------------------------------------------------------------
# NOTE: unlike the other checks in this repo (where green "yes" = good), a
# *match* here is a negative signal, so the color scheme is inverted:
#   on_list  red    — venue matches Beall's List by a high-confidence signal
#   review   orange — fuzzy / near match; a human should verify
#   clean    green  — venue identified and not on the list
#   no_venue grey   — preprint or no venue info; not classifiable
#   error    dark   — record could not be processed
STATUS_FILL_COLORS = {
    "on_list":  "FFC7CE",
    "review":   "FFD79A",
    "clean":    "C6EFCE",
    "no_venue": "D9D9D9",
    "error":    "C00000",
}
