"""Match a single paper's publication venue against the Beall's List snapshot.

This is the heart of the check.  Given the vendored snapshot and one Semantic
Scholar paper record, it decides one of:

  on_list   — the venue matches a Beall entry by a high-confidence signal
              (exact domain, exact ISSN, or exact normalized name)
  review    — only a softer signal matched (open-access-PDF domain, same
              registrable domain, or a fuzzy name match); a human should verify
  clean     — the venue was identified and did not match any entry
  no_venue  — preprint server or no venue info; not classifiable at all

Design choices that keep results *truthful*:
  * Signals are tried strongest-first and the first hit wins, so every row's
    ``matched_on`` says exactly *why* it was flagged.
  * Domain matching is preferred over name matching — it is far less ambiguous
    than fuzzy publisher names.
  * Preprint detection looks only at the *venue* domain/name, never at the
    open-access PDF host (published papers routinely have an arXiv PDF).
  * Nothing is asserted as ``on_list`` from a fuzzy match; fuzzy only ever
    yields ``review``.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from normalize import (
    normalize_name, normalize_host, registrable_domain, normalize_issn,
    extract_issns, significant_tokens, similarity,
)
from config import (
    FUZZY_CUTOFF, MIN_NAME_LEN, FUZZY_LIST_SOURCES,
    PREPRINT_DOMAINS, PREPRINT_NAMES,
)

# When several entries share a domain, prefer the more paper-relevant list.
_LIST_PRIORITY = {
    "publishers": 0, "standalone_journal": 1, "hijacked": 2,
    "vanity_press": 3, "misleading_metric": 4,
}


def _best(entries):
    """Pick the most paper-relevant entry from a list of candidates."""
    return min(entries, key=lambda e: _LIST_PRIORITY.get(e["list_source"], 9))


class BeallIndex:
    """Indexes the snapshot entries for O(1) exact lookups + blocked fuzzy."""

    def __init__(self, entries):
        self.entries = entries
        self.by_domain = {}
        self.by_regdomain = {}
        self.by_name = {}
        self.by_issn = {}
        self.by_token = {}      # significant token -> set of entry indices

        for idx, e in enumerate(entries):
            dom = e.get("domain", "")
            if dom:
                self.by_domain.setdefault(dom, []).append(e)
                self.by_regdomain.setdefault(registrable_domain(dom), []).append(e)
            nm = e.get("normalized_name", "")
            if nm:
                self.by_name.setdefault(nm, []).append(e)
            for issn in e.get("issns", []):
                self.by_issn.setdefault(normalize_issn(issn), []).append(e)
            if e["list_source"] in FUZZY_LIST_SOURCES:
                for tok in significant_tokens(nm):
                    self.by_token.setdefault(tok, set()).add(idx)

    # -- fuzzy ------------------------------------------------------------
    def fuzzy_best(self, normalized_candidate):
        """Best (entry, score) for a candidate name via token-blocked difflib."""
        toks = significant_tokens(normalized_candidate)
        if not toks:
            return None, 0.0
        cand_idxs = set()
        for t in toks:
            cand_idxs |= self.by_token.get(t, set())
        best_entry, best_score = None, 0.0
        for i in cand_idxs:
            score = similarity(normalized_candidate, self.entries[i]["normalized_name"])
            if score > best_score:
                best_entry, best_score = self.entries[i], score
        return best_entry, best_score


def _blank_result(paper, source_file):
    """Build the result skeleton with the venue fields echoed for auditing."""
    ext = paper.get("externalIds") or {}
    pv = paper.get("publicationVenue") or {}
    return {
        "source_file": source_file,
        "paperId": paper.get("paperId", ""),
        "title": paper.get("title", "") or "",
        "year": paper.get("year", "") or "",
        "doi": ext.get("DOI", "") or "",
        "venue": paper.get("venue", "") or "",
        "resolved_name": pv.get("name") or paper.get("venue", "") or "",
        "venue_type": pv.get("type", "") or "",
        "venue_domain": "",
        "issn": normalize_issn(pv.get("issn")) if pv.get("issn") else "",
        "url": paper.get("url", "") or "",
        # filled in by matching:
        "status": "clean",
        "list_source": "",
        "matched_name": "",
        "matched_url": "",
        "matched_on": "",
        "confidence": "",
        "fuzzy_score": "",
        "note": "",
    }


def match_paper(index: BeallIndex, paper: dict, source_file: str) -> dict:
    """Classify one paper. Returns a fully-populated result dict."""
    r = _blank_result(paper, source_file)
    pv = paper.get("publicationVenue") or {}

    # ── candidate signals ────────────────────────────────────────────────
    primary_urls = [pv.get("url")] + list(pv.get("alternate_urls") or [])
    primary_domains = [normalize_host(u) for u in primary_urls if u]
    primary_domains = [d for d in primary_domains if d]
    if primary_domains:
        r["venue_domain"] = primary_domains[0]

    oa = paper.get("openAccessPdf") or {}
    oa_domain = normalize_host(oa.get("url")) if oa.get("url") else ""

    name_sources = [paper.get("venue"), pv.get("name")] + list(
        pv.get("alternate_names") or []
    )
    norm_names, seen = [], set()
    for n in name_sources:
        nn = normalize_name(n)
        if nn and nn not in seen:
            seen.add(nn)
            norm_names.append(nn)

    issns = extract_issns(pv.get("issn"))

    # ── no-venue / preprint classification (venue signals only) ──────────
    is_preprint = any(d in PREPRINT_DOMAINS for d in primary_domains) or (
        not primary_domains and any(n in PREPRINT_NAMES for n in norm_names)
    )
    if is_preprint:
        r["status"] = "no_venue"
        r["note"] = "Preprint server — no journal/publisher venue to classify"
        return r
    if not primary_domains and not norm_names and not issns:
        r["status"] = "no_venue"
        r["note"] = "No venue information in record"
        return r

    # ── high-confidence signals (asserted as on_list) ────────────────────
    for dom in primary_domains:
        if dom in index.by_domain:
            return _hit(r, _best(index.by_domain[dom]), "on_list", "high",
                        "domain", dom)

    for issn in issns:
        if issn in index.by_issn:
            return _hit(r, _best(index.by_issn[issn]), "on_list", "high",
                        "issn", issn)

    for nm in norm_names:
        if len(nm) >= MIN_NAME_LEN and nm in index.by_name:
            return _hit(r, _best(index.by_name[nm]), "on_list", "high",
                        "name_exact", nm)

    # ── softer signals (flagged for review only) ─────────────────────────
    if oa_domain and oa_domain in index.by_domain:
        return _hit(r, _best(index.by_domain[oa_domain]), "review", "medium",
                    "domain_open_access_pdf", oa_domain)

    for dom in primary_domains:
        reg = registrable_domain(dom)
        if reg and reg != dom and reg in index.by_regdomain:
            return _hit(r, _best(index.by_regdomain[reg]), "review", "medium",
                        "domain_registrable", reg)

    best_entry, best_score = None, 0.0
    for nm in norm_names:
        e, s = index.fuzzy_best(nm)
        if s > best_score:
            best_entry, best_score = e, s
    if best_entry and best_score >= FUZZY_CUTOFF:
        r = _hit(r, best_entry, "review", "medium", "name_fuzzy",
                 best_entry["name"])
        r["fuzzy_score"] = round(best_score, 3)
        return r

    # ── nothing matched ──────────────────────────────────────────────────
    if r["venue_type"] == "conference":
        r["note"] = "Conference venue — outside Beall's journal/publisher scope"
    return r


def _hit(r, entry, status, confidence, matched_on, matched_value):
    """Populate the match fields of a result dict from a snapshot entry."""
    r["status"] = status
    r["list_source"] = entry["list_source"]
    r["matched_name"] = entry["name"]
    r["matched_url"] = entry["url"]
    r["matched_on"] = matched_on
    r["confidence"] = confidence
    if not r["note"]:
        r["note"] = f"Matched Beall '{entry['list_source']}' on {matched_on}: {matched_value}"
    return r
