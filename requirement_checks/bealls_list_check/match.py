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
    normalize_name, normalize_host, normalize_issn,
    extract_issns, significant_tokens, similarity,
)
from config import (
    FUZZY_CUTOFF, MIN_NAME_LEN, FUZZY_LIST_SOURCES, WEAK_LIST_SOURCES,
    PREPRINT_DOMAINS, PREPRINT_NAMES, GENERIC_HOSTS,
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
        self.by_name = {}
        self.by_issn = {}
        self.by_token = {}      # significant token -> set of entry indices

        for idx, e in enumerate(entries):
            dom = e.get("domain", "")
            if dom and dom not in GENERIC_HOSTS:
                self.by_domain.setdefault(dom, []).append(e)
            nm = e.get("normalized_name", "")
            if nm:
                self.by_name.setdefault(nm, []).append(e)
            for issn in e.get("issns", []):
                self.by_issn.setdefault(normalize_issn(issn), []).append(e)
            if e["list_source"] in FUZZY_LIST_SOURCES:
                for tok in significant_tokens(nm):
                    self.by_token.setdefault(tok, set()).add(idx)

    # -- domain (exact host, or paper host is a subdomain of a listed host) --
    def domain_match(self, host):
        """Match a paper host against listed hosts, longest suffix first.

        Returns ``(entry, matched_on)`` where ``matched_on`` is ``"domain"``
        for an exact host match or ``"domain_subdomain"`` when the paper host
        is a subdomain of a listed host (e.g. ``journal.mdpi.com`` -> listed
        ``mdpi.com``).

        Crucially, only *actual listed hosts* are ever used as keys, so a paper
        host like ``jurnal.unimed.ac.id`` can never collapse onto a public
        suffix such as ``ac.id`` — that whole class of false positives is gone
        by construction.
        """
        if not host:
            return None, None
        labels = host.split(".")
        for i in range(len(labels) - 1):
            cand = ".".join(labels[i:])
            if cand in self.by_domain:
                return _best(self.by_domain[cand]), (
                    "domain" if i == 0 else "domain_subdomain"
                )
        return None, None

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


def _paper_url(paper, ext):
    """Best link to the paper *at its publisher*.

    Prefers the DOI (which resolves to the publisher's landing page for the
    article), then the open-access PDF, then the Semantic Scholar page.
    """
    doi = ext.get("DOI") or ""
    if doi:
        return f"https://doi.org/{doi}"
    oa = paper.get("openAccessPdf") or {}
    return oa.get("url") or paper.get("url") or ""


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
        "paper_url": _paper_url(paper, ext),
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
    }


def match_paper(index: BeallIndex, paper: dict, source_file: str) -> dict:
    """Classify one paper. Returns a fully-populated result dict."""
    r = _blank_result(paper, source_file)
    pv = paper.get("publicationVenue") or {}

    # ── candidate signals ────────────────────────────────────────────────
    # Semantic Scholar sometimes MERGES several same-named journals into one
    # venue record, so alternate_urls can point at unrelated publishers (a
    # legit Elsevier journal + a predatory clone, say).  We therefore separate
    # the canonical venue URL (publicationVenue.url) from the alternate URLs and
    # trust them differently: a canonical-domain match is asserted (on_list),
    # but an alternate-only domain match is merely flagged for review.
    canonical_domain = normalize_host(pv.get("url")) if pv.get("url") else ""
    alt_domains = [normalize_host(u) for u in (pv.get("alternate_urls") or []) if u]
    alt_domains = [d for d in alt_domains if d]
    all_domains = ([canonical_domain] if canonical_domain else []) + alt_domains
    r["venue_domain"] = canonical_domain or (alt_domains[0] if alt_domains else "")

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
    is_preprint = any(d in PREPRINT_DOMAINS for d in all_domains) or (
        not all_domains and any(n in PREPRINT_NAMES for n in norm_names)
    )
    if is_preprint:
        r["status"] = "no_venue"
        return r
    if not all_domains and not norm_names and not issns:
        r["status"] = "no_venue"
        return r

    # ── high-confidence signals (asserted as on_list) ────────────────────
    # Canonical venue domain: exact host or subdomain of a listed host
    # (e.g. journal.mdpi.com -> mdpi.com).
    if canonical_domain:
        entry, matched_on = index.domain_match(canonical_domain)
        if entry:
            return _hit(r, entry, "on_list", matched_on)

    for issn in issns:
        if issn in index.by_issn:
            return _hit(r, _best(index.by_issn[issn]), "on_list", "issn")

    for nm in norm_names:
        # Require >= 2 words: single generic words ("Research", "Heliyon")
        # collide with legitimate journals and cause false positives.
        if len(nm) >= MIN_NAME_LEN and len(nm.split()) >= 2 and nm in index.by_name:
            return _hit(r, _best(index.by_name[nm]), "on_list", "name_exact")

    # ── softer signals (flagged for review only) ─────────────────────────
    # An ALTERNATE venue URL points at a listed domain.  Because S2 merges
    # same-named journals, this may be a different journal than the paper's —
    # surface it but do not assert it.
    for dom in alt_domains:
        entry, _ = index.domain_match(dom)
        if entry:
            return _hit(r, entry, "review", "domain_alternate_url")

    # The open-access PDF is hosted on a listed publisher domain, but the
    # venue itself was not identified as such — worth a human glance.
    if oa_domain:
        entry, _ = index.domain_match(oa_domain)
        if entry:
            return _hit(r, entry, "review", "domain_open_access_pdf")

    best_entry, best_score = None, 0.0
    for nm in norm_names:
        e, s = index.fuzzy_best(nm)
        if s > best_score:
            best_entry, best_score = e, s
    if best_entry and best_score >= FUZZY_CUTOFF:
        return _hit(r, best_entry, "review", "name_fuzzy")

    # ── nothing matched ──────────────────────────────────────────────────
    return r


def _hit(r, entry, status, matched_on):
    """Populate the match fields of a result dict from a snapshot entry.

    Matches against the "weak" lists (vanity-press book publishers, fake-metric
    companies) are downgraded from on_list -> review: they are surfaced for a
    human but never asserted as a high-confidence predatory verdict.
    """
    if status == "on_list" and entry["list_source"] in WEAK_LIST_SOURCES:
        status = "review"
    r["status"] = status
    r["list_source"] = entry["list_source"]
    r["matched_name"] = entry["name"]
    r["matched_url"] = entry["url"]
    r["matched_on"] = matched_on
    return r
