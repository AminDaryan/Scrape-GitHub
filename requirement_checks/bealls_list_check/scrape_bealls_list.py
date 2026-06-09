"""Scrape beallslist.net and vendor a local snapshot of every list.

Beall's List is a list of *publishers* and *journals*, not papers.  To decide
whether a paper sits on the list we first need a clean, structured copy of the
list itself.  This script fetches every list page on the (unofficial, archived)
beallslist.net site, parses each entry into ``{name, url, domain, ...}`` and
writes them all to ``data/bealls_snapshot.json``.

Why vendor a snapshot instead of scraping live on every run?
  * Reproducibility — a run today and a run next month compare against the
    same list, so result differences come from the papers, not the website.
  * Resilience — the matcher works offline and does not depend on the site
    being up (it has gone down before).
  * Honesty — the snapshot records *when* it was taken, which matters because
    Beall's List is frozen/contested and not authoritative ground truth.

Pages scraped (all five "list" pages; the informational pages are skipped):
  * /                      predatory *publishers*        (name + URL list)
  * /standalone-journals   standalone predatory journals (name + URL list)
  * /hijacked-journals     hijacked journals             (HTML table)
  * /misleading-metrics    fake-metric companies         (name + URL list)
  * /vanity-press          vanity presses                (name + URL list)

Note on scope: a paper venue realistically only ever matches the publishers,
standalone-journals or hijacked lists.  misleading-metrics (metric vendors) and
vanity-press (book publishers) are included because the user asked for
"everything on the site", but each entry is tagged with ``list_source`` so a
match against those low-relevance lists is visible and filterable.

For the hijacked table only the *hijacker/predatory* column is taken — the
"authentic journal" column lists legitimate journals, and flagging those would
be a false positive.

Usage:
  python scrape_bealls_list.py
"""

import json
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
from normalize import normalize_name, normalize_host, extract_issns

# Make name printing safe on Windows consoles (cp1252 can't encode some names).
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE = "https://beallslist.net"
USER_AGENT = "Mozilla/5.0 (compatible; bealls-list-check/1.0)"
SNAPSHOT_PATH = Path(__file__).resolve().parent / "data" / "bealls_snapshot.json"

# (path, list_source, parse_mode)
LIST_PAGES = [
    ("/",                    "publishers",         "links"),
    ("/standalone-journals", "standalone_journal", "links"),
    ("/hijacked-journals",   "hijacked",           "hijacked_table"),
    ("/misleading-metrics",  "misleading_metric",  "links"),
    ("/vanity-press",        "vanity_press",       "links"),
]

# Anchor texts that are page chrome, not list entries.
_CHROME_TEXTS = {
    "", "home", "contact", "about", "skip to content", "read more",
    "comments", "reply", "log in", "search", "menu",
}

_TAG_RE = re.compile(r"<[^>]+>")
_LI_A_RE = re.compile(
    r"<li[^>]*>\s*<a[^>]+href=\"(https?://[^\"]+)\"[^>]*>(.*?)</a>",
    re.IGNORECASE | re.DOTALL,
)
_TR_RE = re.compile(r"<tr[^>]*>(.*?)</tr>", re.IGNORECASE | re.DOTALL)
_TD_RE = re.compile(r"<t[dh][^>]*>(.*?)</t[dh]>", re.IGNORECASE | re.DOTALL)
_HREF_RE = re.compile(r"href=\"(https?://[^\"]+)\"", re.IGNORECASE)
_HEADING_RE = re.compile(r"<(h[1-4])[^>]*>(.*?)</\1>", re.IGNORECASE | re.DOTALL)


def _text(html_fragment: str) -> str:
    """Strip tags and collapse whitespace from an HTML fragment."""
    return re.sub(r"\s+", " ", _TAG_RE.sub("", html_fragment)).strip()


def strip_excluded_sections(html: str) -> str:
    """Remove any ``<hN>Excluded …</hN>`` section (heading + content up to the
    next heading) from a page before its entries are parsed.

    beallslist.net's Publishers page ends with a section titled
    "Excluded – decide after reading" that lists publishers the author
    deliberately did NOT put on the list (e.g. MDPI: "I decided not to include
    MDPI on the list itself"). Those are explicitly *not* listed, so scraping
    them as entries is wrong. Only this section is removed; the real
    "Original list" and "Update" sections above it are untouched.
    """
    headings = list(_HEADING_RE.finditer(html))
    cuts = []
    for i, m in enumerate(headings):
        if "exclud" in _text(m.group(2)).lower():
            end = headings[i + 1].start() if i + 1 < len(headings) else len(html)
            cuts.append((m.start(), end))
    for start, end in reversed(cuts):      # back-to-front keeps earlier indices valid
        html = html[:start] + html[end:]
    return html


def fetch(path: str) -> str:
    """GET a page and return its decoded HTML (raises on HTTP error)."""
    url = BASE + path
    resp = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=30)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_links(html: str, list_source: str) -> list:
    """Parse a ``<li><a href=...>Name</a>`` list page into entries.

    Keeps only external links (the entry websites); the site's own nav/footer
    links point back to beallslist.net and are dropped.
    """
    entries = []
    for url, inner in _LI_A_RE.findall(html):
        if "beallslist.net" in url.lower():
            continue
        name = _text(inner)
        if name.lower() in _CHROME_TEXTS or len(name) < 2:
            continue
        entries.append(_make_entry(name, url, list_source))
    return entries


def parse_hijacked_table(html: str, list_source: str) -> list:
    """Parse the hijacked-journals table, taking only the predatory column.

    Row layout is ``[Hijacker/Predatory Journal, Authentic Journal]``.  We take
    the first cell (the fake journal) plus any URL inside it; the authentic
    journal column is deliberately ignored so legitimate journals are never
    flagged.
    """
    entries = []
    for row in _TR_RE.findall(html):
        cells = _TD_RE.findall(row)
        if not cells:
            continue
        hijacker_html = cells[0]
        name = _text(hijacker_html)
        if not name or name.lower() in {"hijacker/predatory journal",
                                        "hijacked/predatory journal",
                                        "predatory journal"}:
            continue
        href_match = _HREF_RE.search(hijacker_html)
        url = href_match.group(1) if href_match else ""
        # ISSNs sometimes appear in the row text; capture from the whole row.
        issns = extract_issns(_text(row))
        entries.append(_make_entry(name, url, list_source, issns=issns))
    return entries


def _make_entry(name: str, url: str, list_source: str, issns=None) -> dict:
    """Build a normalized snapshot entry."""
    host = normalize_host(url)
    return {
        "name": name,
        "url": url,
        "domain": host,
        "normalized_name": normalize_name(name),
        "issns": issns or [],
        "list_source": list_source,
    }


def scrape() -> dict:
    """Scrape all list pages and return the snapshot dict."""
    all_entries = []
    counts = {}
    for path, list_source, mode in LIST_PAGES:
        print(f"Fetching {BASE + path} ...", end=" ", flush=True)
        html = strip_excluded_sections(fetch(path))
        if mode == "hijacked_table":
            entries = parse_hijacked_table(html, list_source)
        else:
            entries = parse_links(html, list_source)
        print(f"{len(entries)} entries")
        counts[list_source] = len(entries)
        all_entries.extend(entries)
        time.sleep(0.5)

    # Deduplicate within the combined list by (list_source, normalized_name, domain).
    seen, deduped = set(), []
    for e in all_entries:
        key = (e["list_source"], e["normalized_name"], e["domain"])
        if key in seen:
            continue
        seen.add(key)
        deduped.append(e)

    return {
        "meta": {
            "source": BASE,
            "scraped_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "pages": [p for p, _, _ in LIST_PAGES],
            "counts_raw": counts,
            "total_entries": len(deduped),
            "note": (
                "Unofficial archived copy of Beall's List. Appearance on this "
                "list is an allegation as of the scrape date, not authoritative "
                "proof a venue is predatory. The list is frozen/contested."
            ),
        },
        "entries": deduped,
    }


def main():
    snapshot = scrape()
    SNAPSHOT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(SNAPSHOT_PATH, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=1)

    meta = snapshot["meta"]
    print("-" * 60)
    print(f"Scraped {meta['total_entries']} unique entries:")
    for src, n in meta["counts_raw"].items():
        print(f"  {src:20} {n}")
    with_domain = sum(1 for e in snapshot["entries"] if e["domain"])
    print(f"Entries with a usable domain: {with_domain}")
    print(f"Saved snapshot -> {SNAPSHOT_PATH}")


if __name__ == "__main__":
    main()
