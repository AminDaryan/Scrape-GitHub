"""Shared normalization + fuzzy-matching helpers for the Beall's List check.

Both the scraper (which builds the vendored Beall's List snapshot) and the
matcher (which compares each paper's venue against that snapshot) need to
compare publisher / journal names and web domains in a consistent way.  Doing
the normalization in one place is what makes a match meaningful: "MDPI AG",
"mdpi" and "https://www.mdpi.com/" must all collapse to the same key.

Three kinds of keys are produced:

  * normalized name  — lower-cased, accent-stripped, punctuation-removed,
                       whitespace-collapsed, leading "the " dropped.  Used for
                       exact-name matching and as the input to fuzzy matching.
  * normalized host  — the registrable web host of a URL, lower-cased and with
                       a leading "www." removed.  Used for domain matching,
                       which is the single most reliable signal we have.
  * normalized ISSN  — 8 chars as ``NNNN-NNNX``.  Used for journal matching
                       when both sides carry an ISSN.

The fuzzy tier is intentionally dependency-free (``difflib``) and uses token
blocking so we never compare every paper against every Beall entry.
"""

import re
import unicodedata
from difflib import SequenceMatcher
from urllib.parse import urlparse

# ---------------------------------------------------------------------------
# Name normalization
# ---------------------------------------------------------------------------

_WHITESPACE_RE = re.compile(r"\s+")
_NON_ALNUM_RE = re.compile(r"[^a-z0-9]+")


def strip_accents(text: str) -> str:
    """Remove combining diacritics so "Málaga" compares equal to "Malaga"."""
    decomposed = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def normalize_name(name) -> str:
    """Collapse a publisher/journal name to a comparison key.

    Lower-cases, strips accents, turns ``&`` into ``and``, removes all
    punctuation, collapses whitespace and drops a leading ``the``.  Returns an
    empty string for falsy input.
    """
    if not name:
        return ""
    text = strip_accents(str(name)).lower()
    text = text.replace("&", " and ")
    text = _NON_ALNUM_RE.sub(" ", text)
    text = _WHITESPACE_RE.sub(" ", text).strip()
    if text.startswith("the "):
        text = text[4:]
    return text


# ---------------------------------------------------------------------------
# Host / domain normalization
# ---------------------------------------------------------------------------

# Common multi-label public suffixes, so "journal.co.uk" yields "journal.co.uk"
# (its registrable domain) rather than the meaningless "co.uk".
_TWO_LEVEL_SUFFIXES = {
    "co.uk", "ac.uk", "org.uk", "gov.uk",
    "com.au", "org.au", "edu.au", "gov.au",
    "co.in", "ac.in", "org.in", "edu.in", "gov.in",
    "co.jp", "or.jp", "ac.jp",
    "com.br", "org.br", "gov.br",
    "co.za", "org.za",
    "com.cn", "org.cn", "edu.cn",
    "com.tr", "org.tr",
    "com.pk", "edu.pk",
}


def normalize_host(url_or_host) -> str:
    """Return the lower-cased host of a URL with any leading ``www.`` removed.

    Accepts either a full URL ("http://www.mdpi.com/journal/x") or a bare host
    ("www.mdpi.com").  Returns an empty string when nothing host-like is found.
    """
    if not url_or_host:
        return ""
    raw = str(url_or_host).strip()
    if "//" not in raw:
        raw = "//" + raw
    host = urlparse(raw).netloc.lower()
    host = host.split("@")[-1]      # drop any user:pass@
    host = host.split(":")[0]       # drop any :port
    if host.startswith("www."):
        host = host[4:]
    return host


def registrable_domain(host: str) -> str:
    """Best-effort registrable domain (eTLD+1) of a normalized host.

    This is a pragmatic approximation, not a full Public Suffix List lookup:
    it handles a curated set of common two-label suffixes (``co.uk`` etc.) and
    otherwise takes the last two labels.  Good enough to treat
    ``journal.mdpi.com`` and ``www.mdpi.com`` as the same publisher while
    keeping the dependency surface at zero.
    """
    if not host:
        return ""
    parts = host.split(".")
    if len(parts) <= 2:
        return host
    last_two = ".".join(parts[-2:])
    if last_two in _TWO_LEVEL_SUFFIXES:
        return ".".join(parts[-3:])
    return last_two


# ---------------------------------------------------------------------------
# ISSN normalization
# ---------------------------------------------------------------------------

_ISSN_RE = re.compile(r"\b(\d{4})-?(\d{3}[\dxX])\b")


def normalize_issn(issn) -> str:
    """Return an ISSN as ``NNNN-NNNX`` (upper-cased check digit), or ""."""
    if not issn:
        return ""
    match = _ISSN_RE.search(str(issn))
    if not match:
        return ""
    return f"{match.group(1)}-{match.group(2).upper()}"


def extract_issns(text) -> list:
    """Pull every ISSN-looking token out of free text, normalized + deduped."""
    if not text:
        return []
    out, seen = [], set()
    for m in _ISSN_RE.finditer(str(text)):
        issn = f"{m.group(1)}-{m.group(2).upper()}"
        if issn not in seen:
            seen.add(issn)
            out.append(issn)
    return out


# ---------------------------------------------------------------------------
# Fuzzy matching (dependency-free, token-blocked)
# ---------------------------------------------------------------------------

# Words too generic to block on: present in thousands of venue names, so they
# would defeat the purpose of blocking (and drive false positives).
_FUZZY_STOPWORDS = {
    "journal", "journals", "international", "national", "of", "the", "and",
    "for", "in", "on", "review", "reviews", "research", "science", "sciences",
    "studies", "study", "press", "publishing", "publisher", "publishers",
    "publication", "publications", "open", "access", "university", "college",
    "academic", "academy", "global", "world", "advanced", "advances",
    "current", "modern", "applied", "general", "society", "association",
    "institute", "education", "technology", "engineering", "medical",
    "medicine", "health", "scientific", "letters", "annals", "archives",
    "bulletin", "proceedings", "conference", "transactions",
}


def significant_tokens(normalized_name: str) -> set:
    """Tokens worth blocking on: length >= 4 and not a generic stopword."""
    return {
        tok for tok in normalized_name.split()
        if len(tok) >= 4 and tok not in _FUZZY_STOPWORDS
    }


def similarity(a: str, b: str) -> float:
    """Symmetric 0..1 similarity between two normalized names (difflib ratio)."""
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()
