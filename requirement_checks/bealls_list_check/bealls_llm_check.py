"""Optional LLM second pass over the Beall's List check (recall booster).

The deterministic matcher in ``match.py`` is precise but structurally misses
cases it cannot see: a journal whose Semantic Scholar URL is on an *alias*
domain (e.g. a paper in *WSEAS Transactions* whose S2 URL is ``worldses.org``
while Beall lists the publisher under ``wseas.org``), or where the list records
a *publisher* name but the paper carries the *journal* name.  An LLM with world
knowledge closes that gap.

What this module does
---------------------
1. Run the deterministic check (``match.py``) over the whole corpus.
2. Collect the DISTINCT venues (deduplicated by normalized name + domain +
   ISSN), so the LLM is asked once per venue, not once per paper.
3. For every distinct venue, ask the LLM — grounded with (a) the nearest
   Beall's List entries, (b) Beall's "how to recognize predatory journals"
   criteria, and (c) its own knowledge — whether the venue or its publisher is
   on Beall's List / clearly predatory.
4. Fold the verdict back in:
     * a venue currently ``clean``/``no_venue`` that the LLM judges predatory
       is promoted to ``review`` (a human still confirms — we never assert
       ``on_list`` from the LLM, so recall rises without blind trust);
     * every checked paper gets an "LLM assessment" column recording the
       verdict + reason, clearly tagged as LLM-decided.
5. Write ``results/bealls_llm_results.xlsx`` (same layout as the deterministic
   workbook, plus the LLM column and an "LLM second pass" token counter in the
   Summary).

This is OPT-IN and costs LLM tokens.  It uses whatever provider/deployment your
``.env`` configures (``LLM_PROVIDER`` / ``AZURE_OPENAI_DEPLOYMENT``).

Usage:
  python bealls_llm_check.py            # all distinct venues
  python bealls_llm_check.py --limit 50 # first 50 distinct venues (cheap test)
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

import bealls_list_check as det                       # the deterministic pass + Excel writer
from match import BeallIndex, match_paper
from normalize import normalize_name, significant_tokens, similarity

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
    """Assess every distinct venue with the LLM and fold verdicts into *results*.

    Returns an ``llm_stats`` dict for the Summary sheet.
    """
    # Group result rows by venue; skip venues with no identifying info at all.
    groups = {}
    for r in results:
        # Remember the DETERMINISTIC verdict before we possibly promote it, so
        # the disagreement sheet can contrast deterministic vs LLM.
        r["det_status"] = r["status"]
        key = _venue_key(r)
        if not any(key):
            continue
        groups.setdefault(key, []).append(r)

    # Check the actionable venues (already on_list / review) FIRST, so that even
    # a --limit run always annotates the flagged rows (the ones a human reviews).
    def _flagged_first(key):
        return 0 if groups[key][0]["status"] in ("on_list", "review") else 1
    keys = sorted(groups, key=_flagged_first)
    if limit is not None:
        keys = keys[:limit]
    print(f"LLM pass: {len(keys)} distinct venues to check "
          f"(of {len(groups)} total) using {deployment}")

    llm_yes = promoted = 0
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
        if verdict["on_bealls_list"] == "yes":
            llm_yes += 1

        for r in rows:
            r["llm_verdict"] = verdict["on_bealls_list"]
            r["llm_matched"] = verdict["matched_entity"]
            r["llm_reason"] = verdict["reason"]
            # Recall: promote a clean/no_venue venue the LLM flags to review.
            if verdict["on_bealls_list"] == "yes" and r["status"] in ("clean", "no_venue"):
                r["status"] = "review"
                r["matched_name"] = r.get("matched_name") or verdict["matched_entity"] or "(LLM second pass)"
                r["review_reason"] = ("Flagged by the LLM second pass — verify "
                                      "(see the LLM assessment column).")
                promoted += 1
        time.sleep(0.05)

    return {
        "model": deployment,
        "venues_checked": len(keys),
        "llm_yes": llm_yes,
        "promoted": promoted,
    }


# Verdict -> short label for the first LLM-column row.
_VERDICT_LABEL = {
    "yes": "LLM: likely predatory",
    "no": "LLM: not predatory",
    "uncertain": "LLM: uncertain",
}


def llm_subrows(r):
    """Sub-rows for the 'LLM assessment' extra column (blank if not checked)."""
    verdict = r.get("llm_verdict")
    if not verdict:
        return [("", None)]
    rows = [(_VERDICT_LABEL.get(verdict, f"LLM: {verdict}"), None)]
    if r.get("llm_matched"):
        rows.append((f"LLM match: {r['llm_matched']}", None))
    if r.get("llm_reason"):
        rows.append((f"LLM reason: {r['llm_reason']}", None))
    return rows


def disagrees(r):
    """True if the LLM and the deterministic pass reached opposite conclusions.

    Only confident LLM verdicts count. Compares 'did the deterministic pass flag
    it?' (det_status in on_list/review) with 'did the LLM flag it?' (verdict ==
    yes).  These are the rows most worth a human's attention.
    """
    verdict = r.get("llm_verdict")
    if verdict not in ("yes", "no"):
        return False
    det_flagged = r.get("det_status") in ("on_list", "review")
    return det_flagged != (verdict == "yes")


def main(argv=None):
    parser = argparse.ArgumentParser(description="LLM second pass for the Beall's List check.")
    parser.add_argument("--limit", type=int, default=None,
                        help="check only the first N distinct venues (cheap test run)")
    args = parser.parse_args(argv)

    print("=" * 70)
    print("Beall's List check — LLM second pass")
    print("=" * 70)

    snapshot = det.load_snapshot()
    index = BeallIndex(snapshot["entries"])

    print("Running the deterministic pass first ...")
    results, corpus_files = [], set()
    for source_file, paper in det.load_corpus():
        corpus_files.add(source_file)
        results.append(match_paper(index, paper, source_file))

    token_usage = TokenUsageTracker()
    started = time.time()
    llm_stats = run_llm_pass(results, index, client=default_client,
                             deployment=AZURE_OPENAI_DEPLOYMENT,
                             token_usage=token_usage, limit=args.limit)
    print_token_usage_report(token_usage, AZURE_OPENAI_DEPLOYMENT)

    # Rows where the LLM and the deterministic pass disagree (LLM-found first).
    disagreements = sorted(
        (r for r in results if disagrees(r)),
        key=lambda r: 0 if r.get("det_status") in ("clean", "no_venue") else 1,
    )
    llm_stats["disagreements"] = len(disagreements)

    out_path, n_flagged = det.save_workbook(
        results, snapshot["meta"], sorted(corpus_files),
        extra_columns=[("LLM assessment", 64, llm_subrows)],
        token_usage=token_usage,
        llm_stats=llm_stats,
        disagreement_results=disagreements,
        results_filename="bealls_llm_results.xlsx",
    )
    print(f"\nDone in {time.time() - started:.1f}s. "
          f"LLM promoted {llm_stats['promoted']} paper(s) to review; "
          f"{len(disagreements)} LLM/deterministic disagreement(s). "
          f"Flagged (on_list + review): {n_flagged}")
    print(f"Saved -> {out_path}")


if __name__ == "__main__":
    main()
