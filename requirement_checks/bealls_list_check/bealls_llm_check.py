"""Optional LLM pass that helps you RESOLVE the 'review' papers.

The deterministic matcher in ``match.py`` flags some papers as ``review`` — a
soft/contested match (fuzzy name, alternate URL, hijacked-name, …) that a human
must judge.  Going through them by hand is slow, so this pass asks an LLM for a
concrete recommendation on each one, grounded with world knowledge and the
nearest Beall's List entries.

What this module does
---------------------
1. Run the deterministic check (``match.py``) over the whole corpus.
2. Take ONLY the ``review`` papers and collect their DISTINCT venues
   (deduplicated by normalized name + domain + ISSN), so the LLM is asked once
   per venue, not once per paper.
3. For every such venue, ask the LLM — grounded with (a) the nearest Beall's
   List entries, (b) Beall's "how to recognize predatory journals" criteria, and
   (c) its own knowledge — whether the venue or its publisher is predatory.
4. Write that verdict into an "LLM review" column on the Flagged-only sheet, as
   a recommendation a human can act on: Predatory (exclude) / Legitimate (keep)
   / Uncertain (check), plus a one-line reason.  The LLM never changes a
   deterministic verdict and never asserts ``on_list``.
5. Write ``results/bealls_llm_results.xlsx`` (same layout as the deterministic
   workbook, plus that one column and an "LLM review" section in the Summary).

This is OPT-IN and costs LLM tokens (but only for the review backlog, which is
small).  It uses whatever provider/deployment your ``.env`` configures
(``LLM_PROVIDER`` / ``AZURE_OPENAI_DEPLOYMENT``).

Usage:
  python bealls_llm_check.py            # all 'review' venues
  python bealls_llm_check.py --limit 50 # first 50 'review' venues (cheap test)
"""

import argparse
import contextlib
import io
import sys
import time
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))                       # config, match, normalize, bealls_list_check
sys.path.insert(0, str(_HERE.parent))                # requirement_checks/ -> common, openai_client

import bealls_list_check as det                       # deterministic pass, index/corpus helpers, Excel writer
from normalize import normalize_name, significant_tokens, similarity
from user_lists import load_user_entries, WhitelistMatcher

from common.llm_helpers import TokenUsageTracker, llm_call_parse_retry, print_token_usage_report
from openai_client import client as default_client, AZURE_OPENAI_DEPLOYMENT


# Condensed Beall "how to recognize predatory journals" criteria (2015), used to
# ground the LLM so it can flag criteria-based concerns, not only literal
# membership.  Kept short on purpose — it is prompt context, not a spec.
BEALL_CRITERIA = (
    "Beall's red flags for predatory venues: (1) fake or misleading metrics / "
    "invented impact factors; (2) aggressive, indiscriminate spam soliciting "
    "papers or editorial-board seats; (3) very fast 'acceptance' with little or "
    "no real peer review; (4) hidden or deceptive article-processing charges; "
    "(5) fake, non-existent, or non-consenting editorial board; (6) hidden or "
    "falsified country/ownership; (7) a name that mimics an established journal "
    "(hijacking) or overstates its scope; (8) missing/fake ISSN or no "
    "retraction/preservation policy; (9) bogus indexing claims; (10) one company "
    "running hundreds of journals across unrelated fields."
)

SYSTEM_PROMPT = (
    "You are an expert on scholarly publishing and predatory journals (Beall's "
    "List). You are given one publication venue, a few candidate entries from a "
    "vendored snapshot of Beall's List, and Beall's criteria. Decide whether "
    "THIS venue, or the publisher that runs it, is on Beall's List or clearly "
    "matches the predatory criteria.\n\n"
    "Be conservative and precise:\n"
    "- Answer 'yes' only if you are confident the venue or its publisher is the "
    "predatory entity. Many predatory names mimic legitimate journals — do NOT "
    "say 'yes' for a well-known legitimate venue that merely shares a name.\n"
    "- Answer 'no' for clearly legitimate/established venues.\n"
    "- Answer 'uncertain' when you cannot tell.\n\n"
    "Reply with ONLY a JSON object: "
    '{"on_bealls_list": "yes|no|uncertain", "matched_entity": "<publisher or '
    'journal name, or empty>", "reason": "<one or two sentences>"}'
)

# Returned when the model gives nothing usable even after a retry.
_EMPTY_VERDICT = {
    "on_bealls_list": "uncertain",
    "matched_entity": "",
    "reason": "LLM returned no usable response.",
}


def gather_candidates(index, venue_name, venue_domain, k=5):
    """Up to *k* nearest Beall's List entries for grounding the prompt.

    Combines an exact/subdomain domain hit (if any) with the closest entries by
    name (token-blocked difflib), so the LLM sees concrete candidates rather
    than guessing from memory alone.
    """
    cands, seen = [], set()

    def _add(entry):
        if entry and entry["name"] not in seen:
            seen.add(entry["name"])
            cands.append(entry)

    if venue_domain:
        entry, _ = index.domain_match(venue_domain)
        _add(entry)

    nn = normalize_name(venue_name)
    idxs = set()
    for tok in significant_tokens(nn):
        idxs |= index.by_token.get(tok, set())
    scored = sorted(
        ((similarity(nn, index.entries[i]["normalized_name"]), i) for i in idxs),
        reverse=True,
    )
    for _score, i in scored:
        _add(index.entries[i])
        if len(cands) >= k:
            break
    return cands[:k]


def _build_user_message_fn(venue_info, candidates):
    """Return a ``build_user_message(content)`` closure for llm_call_parse_retry.

    The helper passes ``content`` (unused here — we have no file blob), so we
    embed the venue + candidates + criteria directly and ignore it.
    """
    cand_lines = "\n".join(
        f"  - {c['name']}  (domain: {c.get('domain') or '—'}; list: {c['list_source']})"
        for c in candidates
    ) or "  (no close candidates found)"

    def build(_content):
        return (
            "VENUE TO CLASSIFY:\n"
            f"  name:   {venue_info['name'] or '—'}\n"
            f"  domain: {venue_info['domain'] or '—'}\n"
            f"  ISSN:   {venue_info['issn'] or '—'}\n"
            f"  (deterministic check said: {venue_info['det_status']})\n\n"
            "NEAREST BEALL'S LIST CANDIDATES (from the vendored snapshot):\n"
            f"{cand_lines}\n\n"
            f"BEALL'S CRITERIA:\n{BEALL_CRITERIA}\n\n"
            "Is this venue (or its publisher) on Beall's List / clearly "
            "predatory? Reply with the JSON object only."
        )

    return build


def assess_venue(venue_info, candidates, *, client, deployment, token_usage):
    """One LLM call for one venue. Returns a normalized verdict dict."""
    # llm_call_parse_retry prints a per-call debug preview; silence it so a
    # several-thousand-venue run stays readable (our own progress line remains).
    with contextlib.redirect_stdout(io.StringIO()):
        verdict = llm_call_parse_retry(
            client=client,
            deployment=deployment,
            system_prompt=SYSTEM_PROMPT,
            build_user_message=_build_user_message_fn(venue_info, candidates),
            content="",
            token_usage=token_usage,
            empty_payload=_EMPTY_VERDICT,
            max_completion_tokens=300,
            preview_chars=0,
        )
    answer = str(verdict.get("on_bealls_list", "uncertain")).strip().lower()
    if answer not in ("yes", "no", "uncertain"):
        answer = "uncertain"
    return {
        "on_bealls_list": answer,
        "matched_entity": str(verdict.get("matched_entity", "") or ""),
        "reason": str(verdict.get("reason", "") or ""),
    }


def _venue_key(r):
    """Identity that collapses papers sharing a venue into one LLM call."""
    return (normalize_name(r.get("venue") or ""),
            (r.get("venue_domain") or "").lower(),
            (r.get("issn") or "").upper())


def run_llm_pass(results, index, *, client, deployment, token_usage, limit=None):
    """Assess each distinct 'review' venue with the LLM and fold verdicts in.

    Only ``review`` papers are considered (those are the ones a human must
    resolve).  Returns an ``llm_stats`` dict for the Summary sheet.
    """
    # Group ONLY the 'review' rows by venue; skip any with no identifying info.
    groups = {}
    for r in results:
        if r["status"] != "review":
            continue
        key = _venue_key(r)
        if not any(key):
            continue
        groups.setdefault(key, []).append(r)

    keys = list(groups)
    if limit is not None:
        keys = keys[:limit]
    print(f"LLM review pass: {len(keys)} distinct 'review' venue(s) to assess "
          f"using {deployment}")

    counts = {"yes": 0, "no": 0, "uncertain": 0}
    for n, key in enumerate(keys, 1):
        rows = groups[key]
        rep = rows[0]
        venue_info = {
            "name": rep.get("venue") or rep.get("matched_name") or "",
            "domain": rep.get("venue_domain") or "",
            "issn": rep.get("issn") or "",
            "det_status": rep.get("status"),
        }
        print(f"  [{n}/{len(keys)}] {venue_info['name'][:60]!r} ...", end=" ", flush=True)
        candidates = gather_candidates(index, venue_info["name"], venue_info["domain"])
        verdict = assess_venue(venue_info, candidates, client=client,
                               deployment=deployment, token_usage=token_usage)
        print(verdict["on_bealls_list"])
        counts[verdict["on_bealls_list"]] = counts.get(verdict["on_bealls_list"], 0) + 1

        for r in rows:
            r["llm_verdict"] = verdict["on_bealls_list"]
            r["llm_matched"] = verdict["matched_entity"]
            r["llm_reason"] = verdict["reason"]
        time.sleep(0.05)

    return {
        "model": deployment,
        "review_venues_checked": len(keys),
        "llm_yes": counts["yes"],
        "llm_no": counts["no"],
        "llm_uncertain": counts["uncertain"],
    }


# LLM verdict -> the actionable recommendation shown in the 'LLM review' column.
_VERDICT_LABEL = {
    "yes": "Predatory — recommend EXCLUDE",
    "no": "Legitimate — recommend KEEP",
    "uncertain": "Uncertain — check by hand",
}


def llm_subrows(r):
    """Sub-rows for the 'LLM review' column (blank unless this is a review paper).

    Verdict + (matched entity) + reason, so a reviewer can resolve the paper in
    one glance.  Only 'review' rows carry an ``llm_verdict``.
    """
    verdict = r.get("llm_verdict")
    if not verdict:
        return [("", None)]
    rows = [(f"Verdict: {_VERDICT_LABEL.get(verdict, verdict)}", None)]
    if r.get("llm_matched"):
        rows.append((f"Matched: {r['llm_matched']}", None))
    if r.get("llm_reason"):
        rows.append((f"Reason: {r['llm_reason']}", None))
    return rows


def main(argv=None):
    parser = argparse.ArgumentParser(description="LLM second pass for the Beall's List check.")
    parser.add_argument("--limit", type=int, default=None,
                        help="check only the first N distinct venues (cheap test run)")
    parser.add_argument("--whitelist", help="JSON list of venues; only papers in it are checked")
    parser.add_argument("--blacklist", help="JSON list of venues to add to Beall's List as predatory")
    parser.add_argument("--crossref", choices=["off", "flagged", "all"], default="off",
                        help="cross-check DOIs against Crossref (off / flagged / all)")
    args = parser.parse_args(argv)

    print("=" * 70)
    print("Beall's List check — LLM second pass")
    print("=" * 70)

    index, snapshot, n_black = det.build_index(args.blacklist)
    if n_black:
        print(f"Added {n_black} user blacklist entr(ies) as predatory.")
    whitelist = (WhitelistMatcher(load_user_entries(args.whitelist, "whitelist"))
                 if args.whitelist else None)
    if whitelist is not None:
        print(f"Whitelist active — papers outside it are skipped (out_of_scope).")

    print("Running the deterministic pass first ...")
    results, corpus_files, _n_scope = det.classify_corpus(index, whitelist, crossref=args.crossref)

    token_usage = TokenUsageTracker()
    started = time.time()
    llm_stats = run_llm_pass(results, index, client=default_client,
                             deployment=AZURE_OPENAI_DEPLOYMENT,
                             token_usage=token_usage, limit=args.limit)
    print_token_usage_report(token_usage, AZURE_OPENAI_DEPLOYMENT)

    out_path, n_flagged = det.save_workbook(
        results, snapshot["meta"], sorted(corpus_files),
        flagged_extra_columns=[("LLM review", 64, llm_subrows)],
        token_usage=token_usage,
        llm_stats=llm_stats,
        results_filename="bealls_llm_results.xlsx",
    )
    rv = llm_stats["review_venues_checked"]
    print(f"\nDone in {time.time() - started:.1f}s. "
          f"Assessed {rv} 'review' venue(s): "
          f"{llm_stats['llm_yes']} exclude / {llm_stats['llm_no']} keep / "
          f"{llm_stats['llm_uncertain']} uncertain. "
          f"Flagged (on_list + review): {n_flagged}")
    print(f"Saved -> {out_path}")


if __name__ == "__main__":
    main()
