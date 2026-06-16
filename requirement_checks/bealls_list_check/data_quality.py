"""Pre-flight data-quality checks on the Semantic Scholar (S2) records.

S2 metadata is sometimes wrong, and a wrong verdict is worse than a flagged one.
These checks surface *suspect* records so they aren't trusted blindly.  Two
layers, combined by :func:`validate`:

  * :func:`offline_flags` — structural + internal-consistency checks (no network):
    missing venue metadata, malformed DOI/ISSN, encoding artifacts, and the
    "merged venues" red flag (S2 sometimes fuses two same-named journals, so its
    venue URLs span multiple publishers).

  * :func:`crossref_flags` — cross-check the DOI against Crossref, the
    publisher-deposited registration metadata, to catch genuinely wrong venue /
    ISSN.  Network; free; no key.  Bounded by DOI coverage (preprints/books have
    none).

An empty result means "nothing suspicious found" — NOT a guarantee of
correctness (S2 can be wrong in ways no automated check can see).
"""

import re
import sys
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
from normalize import normalize_name, normalize_host, normalize_issn, significant_tokens

_DOI_RE = re.compile(r"^10\.\d{4,9}/\S+$")
_ENTITY_RE = re.compile(r"&[A-Za-z]+;|&#\d+;|�")        # &amp; &#39; or U+FFFD
_CROSSREF_URL = "https://api.crossref.org/works/"
_UA = "bealls-data-quality/1.0 (research tooling)"

_crossref_cache = {}        # doi -> Crossref "message" dict, or None


def _registrable(host):
    """Crude eTLD+1 (last two labels) — enough to tell publishers apart."""
    parts = host.split(".")
    return ".".join(parts[-2:]) if len(parts) >= 2 else host


def offline_flags(paper):
    """Structural + consistency warnings for one S2 record (no network)."""
    flags = []
    pv = paper.get("publicationVenue") or {}
    ext = paper.get("externalIds") or {}
    venue = paper.get("venue") or ""
    pv_name = pv.get("name") or ""
    issn_raw = pv.get("issn") or ""
    doi = (ext.get("DOI") or "").strip()

    if not (venue or pv_name or pv.get("url") or issn_raw):
        flags.append("No venue metadata (no venue, publicationVenue, or ISSN).")
    if doi and not _DOI_RE.match(doi):
        flags.append(f"DOI looks malformed: {doi!r}.")
    if issn_raw and not normalize_issn(issn_raw):
        flags.append(f"ISSN looks malformed: {issn_raw!r}.")
    for label, val in (("venue", venue), ("publicationVenue.name", pv_name)):
        if val and _ENTITY_RE.search(val):
            flags.append(f"Encoding artifacts in {label}: {val!r}.")

    # Merged-venue red flag: canonical + alternate venue URLs span >1 publisher.
    hosts = [normalize_host(pv["url"])] if pv.get("url") else []
    hosts += [normalize_host(u) for u in (pv.get("alternate_urls") or []) if u]
    regs = {_registrable(h) for h in hosts if h}
    if len(regs) >= 2:
        flags.append("S2 may have merged different journals — venue URLs span "
                     f"{len(regs)} publishers ({', '.join(sorted(regs))}).")
    return flags


def _crossref_lookup(doi, *, timeout=10):
    if doi in _crossref_cache:
        return _crossref_cache[doi]
    message = None
    try:
        resp = requests.get(_CROSSREF_URL + requests.utils.quote(doi, safe=""),
                            headers={"User-Agent": _UA}, timeout=timeout)
        if resp.status_code == 200:
            message = resp.json().get("message")
    except (requests.RequestException, ValueError):
        message = None
    _crossref_cache[doi] = message
    return message


def crossref_flags(paper, *, timeout=10):
    """Warnings from comparing S2's venue/ISSN to Crossref's (network)."""
    doi = ((paper.get("externalIds") or {}).get("DOI") or "").strip()
    if not doi or not _DOI_RE.match(doi):
        return []
    msg = _crossref_lookup(doi, timeout=timeout)
    if msg is None:
        return [f"DOI not found in Crossref (or lookup failed): {doi}."]

    pv = paper.get("publicationVenue") or {}
    s2_names = [paper.get("venue"), pv.get("name")] + list(pv.get("alternate_names") or [])
    cr_titles = msg.get("container-title") or []

    flags = []
    s2_tokens = set().union(*(significant_tokens(normalize_name(n)) for n in s2_names)) \
        if s2_names else set()
    cr_tokens = set().union(*(significant_tokens(normalize_name(t)) for t in cr_titles)) \
        if cr_titles else set()
    # Disagree only when neither side shares a meaningful word (tolerates
    # abbreviations like "WSEAS Trans. Inf. Sci." vs the full Crossref title).
    if s2_tokens and cr_tokens and not (s2_tokens & cr_tokens):
        s2_show = next((n for n in s2_names if n), "?")
        flags.append(f"Venue disagrees with Crossref — S2: {s2_show!r}; "
                     f"Crossref: {cr_titles[0]!r}.")

    s2_issn = normalize_issn(pv.get("issn")) if pv.get("issn") else ""
    cr_issns = {normalize_issn(i) for i in (msg.get("ISSN") or []) if normalize_issn(i)}
    if s2_issn and cr_issns and s2_issn not in cr_issns:
        flags.append(f"ISSN disagrees with Crossref — S2: {s2_issn}; "
                     f"Crossref: {sorted(cr_issns)}.")
    return flags


# Note: these venue checks are invoked through the single shared validator,
# common.input_quality.validate_input(), which all checks use.
