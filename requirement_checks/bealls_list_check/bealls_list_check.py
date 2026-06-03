"""Beall's List predatory-venue check — main entry point.

Loads every Semantic Scholar dump in the corpus folder, matches each paper's
publication venue against the vendored Beall's List snapshot, and writes one
combined Excel workbook with three sheets:

  * Results      — one row per paper, colored by status, with full evidence
                   (what matched, on which signal, at what confidence) so any
                   flag can be audited by hand.
  * Summary      — counts per status + per Beall list, and the run metadata
                   (snapshot date, corpus files) needed to interpret them.
  * Flagged only — just the on_list + review rows, so you are not scrolling
                   tens of thousands of clean/preprint rows to find the hits.

Run the scraper first (scrape_bealls_list.py) to produce the snapshot, then:
  python bealls_list_check.py
"""

import json
import sys
import time
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from config import (
    CORPUS_DIR, CORPUS_GLOB, SNAPSHOT_PATH, RESULTS_DIR, RESULTS_FILENAME,
    STATUS_FILL_COLORS,
)
from match import BeallIndex, match_paper


# =============================================================================
# LOADING
# =============================================================================

def load_snapshot():
    """Load the vendored Beall's List snapshot, or exit with guidance."""
    if not SNAPSHOT_PATH.exists():
        sys.exit(
            f"Snapshot not found: {SNAPSHOT_PATH}\n"
            "Run scrape_bealls_list.py first to vendor the Beall's List."
        )
    with open(SNAPSHOT_PATH, encoding="utf-8") as f:
        return json.load(f)


def load_corpus():
    """Yield (source_file, paper) for every record across all corpus files.

    Deduplicates by ``paperId`` across files (the corpus files are disjoint in
    practice, but this guards against accidental overlap).
    """
    files = sorted(CORPUS_DIR.glob(CORPUS_GLOB))
    if not files:
        sys.exit(f"No corpus files matching {CORPUS_GLOB} in {CORPUS_DIR}")
    seen = set()
    for fp in files:
        with open(fp, encoding="utf-8") as f:
            records = json.load(f)
        kept = 0
        for paper in records:
            pid = paper.get("paperId") or paper.get("title")
            if pid in seen:
                continue
            seen.add(pid)
            kept += 1
            yield fp.name, paper
        print(f"  loaded {kept:>6} papers from {fp.name}")


# =============================================================================
# EXCEL OUTPUT
# =============================================================================

# Column header, width, and which result key feeds it.
COLUMNS = [
    ("#",            6,  None),
    ("Status",       11, "status"),
    ("Source File",  20, "source_file"),
    ("Title",        50, "title"),
    ("Year",         8,  "year"),
    ("Venue",        36, "venue"),
    ("Type",         12, "venue_type"),
    ("Venue Domain", 22, "venue_domain"),
    ("ISSN",         12, "issn"),
    ("DOI",          24, "doi"),
    ("Paper URL",    34, "paper_url"),
    ("Beall List",   18, "list_source"),
    ("Matched Entry", 30, "matched_name"),
    ("Matched On",   22, "matched_on"),
    ("Why Review",   52, "review_reason"),
    ("Matched URL",  34, "matched_url"),
]

# Columns rendered as clickable hyperlinks.
LINK_KEYS = {"paper_url", "matched_url"}

_HEADER_FILL = PatternFill("solid", fgColor="2F5496")
_HEADER_FONT = Font(name="Arial", bold=True, size=10, color="FFFFFF")
_DATA_FONT = Font(name="Arial", size=9)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)


def _border():
    edge = Side(style="thin", color="D0D0D0")
    return Border(left=edge, right=edge, top=edge, bottom=edge)


def _write_results_sheet(ws, results):
    """Write the full per-paper Results sheet."""
    border = _border()
    for col, (header, width, _) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font, cell.fill, cell.alignment, cell.border = (
            _HEADER_FONT, _HEADER_FILL, _CENTER, border
        )
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = "A2"

    center_cols = {1, 2, 5, 7, 9, 12, 14}
    for row_idx, r in enumerate(results, 2):
        fill = PatternFill("solid", fgColor=STATUS_FILL_COLORS.get(r["status"], "FFFFFF"))
        for col, (_, _, key) in enumerate(COLUMNS, 1):
            value = (row_idx - 1) if key is None else r.get(key, "")
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.font = _DATA_FONT
            cell.border = border
            cell.alignment = _CENTER if col in center_cols else _LEFT
            if col == 2:               # color the Status cell by status
                cell.fill = fill
                cell.font = Font(name="Arial", size=9, bold=True)
            elif key in LINK_KEYS and value:
                cell.hyperlink = value
                cell.font = Font(name="Arial", size=9, color="0563C1",
                                 underline="single")


def _write_summary_sheet(ws, results, snapshot_meta, corpus_files):
    """Write counts + run metadata so results are interpretable."""
    border = _border()
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 60

    def header(row, a, b):
        for col, val in ((1, a), (2, b)):
            c = ws.cell(row=row, column=col, value=val)
            c.font, c.fill, c.alignment, c.border = (
                _HEADER_FONT, _HEADER_FILL, _CENTER, border
            )

    def kv(row, label, value):
        ws.cell(row=row, column=1, value=label).font = Font(name="Arial", size=10, bold=True)
        c = ws.cell(row=row, column=2, value=value)
        c.font = Font(name="Arial", size=10)
        c.alignment = _LEFT

    total = len(results)
    status_counts = {s: 0 for s in STATUS_FILL_COLORS}
    list_counts = {}
    for r in results:
        status_counts[r["status"]] = status_counts.get(r["status"], 0) + 1
        if r["list_source"]:
            list_counts[r["list_source"]] = list_counts.get(r["list_source"], 0) + 1

    row = 1
    header(row, "Metric", "Value"); row += 1
    kv(row, "Total papers checked", total); row += 1
    kv(row, "On Beall's List (high confidence)", status_counts.get("on_list", 0)); row += 1
    kv(row, "Needs review (soft match)", status_counts.get("review", 0)); row += 1
    kv(row, "Clean (venue identified, not listed)", status_counts.get("clean", 0)); row += 1
    kv(row, "No venue (preprint / no info)", status_counts.get("no_venue", 0)); row += 1
    kv(row, "Errors", status_counts.get("error", 0)); row += 2

    header(row, "Matches by Beall list", "Count"); row += 1
    for src in sorted(list_counts):
        kv(row, f"  {src}", list_counts[src]); row += 1
    row += 1

    header(row, "Run metadata", ""); row += 1
    kv(row, "Snapshot scraped (UTC)", snapshot_meta.get("scraped_utc", "")); row += 1
    kv(row, "Snapshot entries", snapshot_meta.get("total_entries", "")); row += 1
    kv(row, "Snapshot source", snapshot_meta.get("source", "")); row += 1
    kv(row, "Corpus files", ", ".join(corpus_files)); row += 1
    kv(row, "Check run (local)", datetime.now().strftime("%Y-%m-%d %H:%M:%S")); row += 2

    header(row, "Caveat", ""); row += 1
    kv(row, "Interpretation", snapshot_meta.get("note", "")); row += 1
    ws.cell(row=row, column=1,
            value="'on_list' = exact domain/ISSN/name match. 'review' = fuzzy or "
                  "soft match — verify by hand before trusting.").font = Font(
        name="Arial", size=9, italic=True)


def _write_flagged_sheet(ws, results):
    """Write only the on_list + review rows (the actionable subset)."""
    flagged = [r for r in results if r["status"] in ("on_list", "review")]
    # Sort: on_list first, then by list source, then title.
    order = {"on_list": 0, "review": 1}
    flagged.sort(key=lambda r: (order.get(r["status"], 9), r["list_source"], r["title"]))
    _write_results_sheet(ws, flagged)
    return len(flagged)


# Legend-sheet content as one declarative structure.  Each section is a titled
# mini-table of (col-A, col-B, col-C) rows.  headers=None renders the rows as
# full-width bullet lines (the caveats); swatch=True tints column A by status
# to match the Results sheet colors.  Add/edit documentation here only — the
# renderer below is generic.
_LEGEND_TITLE = "Beall's List Check — Legend / Data Dictionary"
_LEGEND_SECTIONS = [
    {
        "title": "STATUS values (and the row color used in Results)",
        "headers": ("Status", "Confidence", "Meaning"),
        "swatch": True,
        "rows": [
            ("on_list", "high", "Venue matched a CORE Beall list (predatory "
             "publishers, standalone journals, or hijacked journals) by a "
             "high-confidence signal: exact domain, subdomain of a listed "
             "domain, exact ISSN, or exact name. Read as \"appears on Beall's "
             "List.\""),
            ("review", "verify", "A softer or contested match: fuzzy name, "
             "open-access PDF host, an alternate venue URL, or a WEAK list "
             "(vanity-press / misleading-metrics). Check by hand."),
            ("clean", "—", "The venue was identified and did NOT match any "
             "Beall entry."),
            ("no_venue", "—", "Preprint server (e.g. arXiv) or no venue "
             "metadata — there is no journal/publisher to classify."),
            ("error", "—", "The record could not be processed; the Python "
             "error is shown in the 'Matched On' column."),
        ],
    },
    {
        "title": "'Matched On' — which signal produced the match",
        "headers": ("Matched On", "Strength", "Meaning"),
        "rows": [
            ("domain", "strongest", "Venue host exactly equals a listed host "
             "(e.g. mdpi.com == mdpi.com)."),
            ("domain_subdomain", "strong", "Venue host is a subdomain of a "
             "listed host (e.g. journal.mdpi.com -> mdpi.com)."),
            ("issn", "strong", "Venue ISSN exactly equals a listed ISSN."),
            ("name_exact", "strong", "Normalized venue name (>= 2 words) "
             "exactly equals a listed name; case/accents/punctuation ignored."),
            ("domain_alternate_url", "weak (review)", "An ALTERNATE venue URL "
             "points at a listed domain. Semantic Scholar merges same-named "
             "journals, so this may be a different journal — verify."),
            ("name_fuzzy", "weak (review)", "Venue name is >= 93% similar to a "
             "listed name. Never asserted as on_list."),
            ("domain_open_access_pdf", "weak (review)", "The open-access PDF "
             "is hosted on a listed domain though the venue was not identified."),
        ],
    },
    {
        "title": "'Beall List' — which list the entry came from",
        "headers": ("Beall List", "Weight", "Meaning"),
        "rows": [
            ("publishers", "CORE", "Predatory publisher (Beall's main list)."),
            ("standalone_journal", "CORE", "Standalone predatory journal."),
            ("hijacked", "CORE", "Hijacked / cloned journal impersonating a "
             "legitimate one (only the fake side is listed)."),
            ("vanity_press", "WEAK -> review", "Book/monograph publisher (e.g. "
             "IGI Global). Not a journal-quality signal; downgraded to review."),
            ("misleading_metric", "WEAK -> review", "Company selling fake "
             "metrics — not a venue. Downgraded to review."),
        ],
    },
    {
        "title": "Columns in the Results / Flagged-only sheets",
        "headers": ("Column", "", "Meaning"),
        "rows": [
            ("Status", "", "Classification of the paper (see Status section)."),
            ("Source File", "", "Which corpus JSON the paper came from."),
            ("Title / Year", "", "Paper title and year (from Semantic Scholar)."),
            ("Venue", "", "The publication venue name from Semantic Scholar."),
            ("Type", "", "journal / conference (conferences are out of scope)."),
            ("Venue Domain", "", "Host of the venue's website (main match key)."),
            ("ISSN / DOI", "", "Venue ISSN and the paper's DOI."),
            ("Paper URL", "", "Link to the article at its publisher (via DOI)."),
            ("Beall List", "", "Which Beall list the match came from."),
            ("Matched Entry", "", "The exact Beall entry that was matched."),
            ("Matched On", "", "Which signal fired (see 'Matched On' section)."),
            ("Why Review", "", "Plain-English reason a row is 'review' rather "
             "than a definite on_list (blank for on_list/clean/no_venue)."),
            ("Matched URL", "", "The listed Beall entry's own URL."),
        ],
    },
    {
        "title": "Caveats — read before interpreting results",
        "headers": None,
        "rows": [
            ("Beall's List is an unofficial, archived, frozen (~2021), CONTESTED "
             "list. A match is an allegation as of the snapshot date, not proof "
             "a venue is predatory.",),
            ("Well-known publishers like MDPI and Frontiers are on the list but "
             "widely considered legitimate. 'on_list' = \"appears on the list,\" "
             "not \"is predatory.\"",),
            ("This check classifies the VENUE, not the paper's content/quality.",),
            ("It uses NO LLM (0 tokens). Every result is deterministic and "
             "auditable from the Matched On / Matched Entry columns.",),
            ("no_venue is mostly arXiv/preprints, which have no journal "
             "publisher and so cannot be on Beall's List.",),
        ],
    },
]


def _write_legend_sheet(ws):
    """Render the Legend tab from the declarative ``_LEGEND_SECTIONS``."""
    border = _border()
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 95
    bold = Font(name="Arial", size=10, bold=True)
    data = Font(name="Arial", size=10)
    section_font = Font(name="Arial", size=12, bold=True, color="2F5496")

    ws.cell(row=1, column=1, value=_LEGEND_TITLE).font = Font(
        name="Arial", size=14, bold=True)
    row = 3

    for sec in _LEGEND_SECTIONS:
        ws.cell(row=row, column=1, value=sec["title"]).font = section_font
        row += 1

        if sec["headers"] is None:                      # caveats: bullet lines
            for (text,) in sec["rows"]:
                cell = ws.cell(row=row, column=1, value="• " + text)
                cell.font, cell.alignment = data, _LEFT
                ws.merge_cells(start_row=row, start_column=1,
                               end_row=row, end_column=3)
                row += 1
            row += 1
            continue

        for col, header in enumerate(sec["headers"], 1):   # header row
            cell = ws.cell(row=row, column=col, value=header)
            cell.font, cell.fill, cell.alignment, cell.border = (
                _HEADER_FONT, _HEADER_FILL, _CENTER, border)
        row += 1

        for cells in sec["rows"]:                          # data rows
            for col, val in enumerate(cells, 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.font = bold if col == 1 else data
                cell.border, cell.alignment = border, _LEFT
                if col == 1 and sec.get("swatch"):
                    cell.fill = PatternFill(
                        "solid", fgColor=STATUS_FILL_COLORS.get(val, "FFFFFF"))
            row += 1
        row += 1


def save_workbook(results, snapshot_meta, corpus_files):
    """Build and save the combined workbook; return its path."""
    wb = openpyxl.Workbook()
    ws_results = wb.active
    ws_results.title = "Results"
    _write_results_sheet(ws_results, results)

    ws_flagged = wb.create_sheet("Flagged only")
    n_flagged = _write_flagged_sheet(ws_flagged, results)

    ws_summary = wb.create_sheet("Summary")
    _write_summary_sheet(ws_summary, results, snapshot_meta, corpus_files)

    ws_legend = wb.create_sheet("Legend")
    _write_legend_sheet(ws_legend)

    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    out_path = RESULTS_DIR / RESULTS_FILENAME
    try:
        wb.save(out_path)
    except PermissionError:
        # The target is almost certainly open in Excel (Windows locks it).
        # Fall back to a timestamped name so a run is never lost.
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = RESULTS_DIR / f"{Path(RESULTS_FILENAME).stem}_{stamp}.xlsx"
        wb.save(out_path)
        print(f"  (default file was locked — saved to {out_path.name} instead)")
    return out_path, n_flagged


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("=" * 70)
    print("Beall's List predatory-venue check")
    print("=" * 70)

    snapshot = load_snapshot()
    index = BeallIndex(snapshot["entries"])
    print(f"Loaded snapshot: {len(snapshot['entries'])} entries "
          f"(scraped {snapshot['meta'].get('scraped_utc', '?')})")

    print(f"Loading corpus from {CORPUS_DIR} ...")
    started = time.time()
    results = []
    corpus_files = set()
    for source_file, paper in load_corpus():
        corpus_files.add(source_file)
        # match_paper never raises — a bad record returns an 'error' result.
        results.append(match_paper(index, paper, source_file))

    elapsed = time.time() - started
    out_path, n_flagged = save_workbook(
        results, snapshot["meta"], sorted(corpus_files)
    )

    # Console summary.
    counts = {s: 0 for s in STATUS_FILL_COLORS}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    print("-" * 70)
    print(f"Checked {len(results)} papers in {elapsed:.1f}s")
    print(f"  on_list : {counts.get('on_list', 0)}")
    print(f"  review  : {counts.get('review', 0)}")
    print(f"  clean   : {counts.get('clean', 0)}")
    print(f"  no_venue: {counts.get('no_venue', 0)}")
    print(f"  error   : {counts.get('error', 0)}")
    print(f"Flagged (on_list + review): {n_flagged}")
    print(f"Saved -> {out_path}")


if __name__ == "__main__":
    main()
