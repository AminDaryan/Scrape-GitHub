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
    ("Venue (S2)",   34, "venue"),
    ("Resolved Venue", 34, "resolved_name"),
    ("Type",         12, "venue_type"),
    ("Venue Domain", 22, "venue_domain"),
    ("ISSN",         12, "issn"),
    ("DOI",          24, "doi"),
    ("Beall List",   18, "list_source"),
    ("Matched Entry", 30, "matched_name"),
    ("Matched On",   22, "matched_on"),
    ("Confidence",   11, "confidence"),
    ("Fuzzy Score",  11, "fuzzy_score"),
    ("Notes",        46, "note"),
    ("Matched URL",  34, "matched_url"),
]

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

    center_cols = {1, 2, 5, 8, 10, 12, 14, 15, 16}
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

    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    out_path = RESULTS_DIR / RESULTS_FILENAME
    wb.save(out_path)
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
        try:
            results.append(match_paper(index, paper, source_file))
        except Exception as exc:  # never let one bad record kill the run
            results.append({
                "source_file": source_file,
                "paperId": paper.get("paperId", ""),
                "title": paper.get("title", "") or "",
                "status": "error", "note": f"{type(exc).__name__}: {exc}",
                **{k: "" for k in (
                    "year", "doi", "venue", "resolved_name", "venue_type",
                    "venue_domain", "issn", "url", "list_source",
                    "matched_name", "matched_url", "matched_on",
                    "confidence", "fuzzy_score")},
            })

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
