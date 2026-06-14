"""User-supplied whitelist / blacklist for the Beall's check.

Two optional inputs, both JSON lists of ``{"name": ..., "domain": ...}`` (a
plain string is treated as a name; ``url`` is accepted in place of ``domain``):

  * **blacklist** — venues the user asserts are predatory.  They are turned into
    snapshot entries (``list_source="blacklist"``) and added to the BeallIndex,
    so a paper in one of them is flagged exactly like a real Beall entry.

  * **whitelist** — a *scope filter*.  When non-empty, only papers whose venue
    matches a whitelist entry are evaluated; every other paper is marked
    ``out_of_scope`` and skipped.  An empty/absent whitelist means "check
    everything" (the default behaviour).

A paper is *in scope* if any of its venue domains is the same as, or a subdomain
of, a whitelisted domain (so whitelisting a publisher domain covers all its
journals), OR a whitelisted name appears as a whole-word phrase in one of the
paper's venue names (so whitelisting the publisher "IEEE" matches "IEEE
Transactions on …", without "nature" spuriously matching "signature").
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from normalize import normalize_name, normalize_host, normalize_issn
from match import venue_signals


def load_user_entries(path, list_source):
    """Load a user JSON list into snapshot-entry dicts (``list_source`` tagged).

    Skips entries that have neither a usable name nor a domain.
    """
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list):
        raise ValueError(f"{path} must be a JSON list (or single object) of venues.")

    entries = []
    for item in data:
        if isinstance(item, str):
            item = {"name": item}
        name = (item.get("name") or "").strip()
        url = (item.get("url") or item.get("domain") or "").strip()
        domain = normalize_host(url) if url else ""
        normalized_name = normalize_name(name)
        if not (domain or normalized_name):
            continue
        entries.append({
            "name": name or domain,
            "url": item.get("url") or (f"https://{domain}" if domain else ""),
            "domain": domain,
            "normalized_name": normalized_name,
            "issns": [normalize_issn(i) for i in (item.get("issns") or []) if normalize_issn(i)],
            "list_source": list_source,
        })
    return entries


class WhitelistMatcher:
    """Decides whether a paper's venue is in the user's whitelist scope."""

    def __init__(self, entries):
        self.domains = {e["domain"] for e in entries if e["domain"]}
        # Stored space-padded so we can test for whole-word phrase containment.
        self.name_phrases = {f" {e['normalized_name']} " for e in entries if e["normalized_name"]}

    @property
    def active(self):
        return bool(self.domains or self.name_phrases)

    def _domain_in_scope(self, host):
        if not host:
            return False
        labels = host.split(".")
        return any(".".join(labels[i:]) in self.domains for i in range(len(labels) - 1))

    def in_scope(self, paper):
        """True if the paper's venue belongs to a whitelisted publisher/journal."""
        domains, names = venue_signals(paper)
        if any(self._domain_in_scope(d) for d in domains):
            return True
        padded = [f" {n} " for n in names]
        return any(phrase in pn for phrase in self.name_phrases for pn in padded)
