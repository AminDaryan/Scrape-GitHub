"""
Question 5.1.3 — Pre-processing & Pipeline Code classifier

Checks ONLY the paper appendix (no repo inspection) for structured preprocessing
pseudocode or template tables.

Logic:
  1. Resolve the paper's arXiv ID by title.                           → fetch_appendix()
  2. Fetch the HTML render (ar5iv first, arxiv.org/html as fallback)
     and reject navigation/stub pages.                                → fetch_appendix(), is_junk()
  3. Slice out the appendix section; fall back to the last 25% of
     the text when no explicit appendix heading is found.             → fetch_appendix()
  4. Send the appendix to the LLM and classify it.                   → process(), call_gpt()
  5. Parse the JSON response, applying heuristic overrides to
     suppress false positives from auxiliary prompts.                 → parse()
  6. When arXiv fetch fails, fall back to GPT's training knowledge.  → process()

Labels:
  SEPARATE_APPENDIX — appendix has structured pseudocode/template table for preprocessing
  MISSING           — everything else
"""

import json

import os, re, sys, time, urllib.parse, requests
from pathlib import Path
from collections import Counter
from dotenv import load_dotenv

load_dotenv()

sys.path.append(str(Path(__file__).resolve().parent.parent))
from openai_client import client, AZURE_OPENAI_DEPLOYMENT
from common.token_usage import TokenUsageTracker, print_token_usage_report
from papers_from_database import PAPERS
from prompts import SYSTEM_PROMPT, USER_PROMPT, FALLBACK_PROMPT

PAUSE   = 1.5
ALLOWED = {"SEPARATE_APPENDIX", "MISSING"}
TOKEN_USAGE = TokenUsageTracker()

# ---------------------------------------------------------------------------
# Prompts
# ---------------------------------------------------------------------------
# Prompt templates are defined in neighboring prompts.py.

# ---------------------------------------------------------------------------
# GPT call
# ---------------------------------------------------------------------------

def call_gpt(messages):
    """Send a list of role-tagged messages (system + user) to the configured Azure OpenAI
    deployment via the chat completions API and return the model's reply as a raw string.

    The messages list follows the OpenAI chat format: each entry has a 'role'
    ('system' or 'user') and a 'content' string. The model generates the next
    message in the conversation — in this case, a JSON classification verdict.
    Tracks token usage for the post-run cost report and logs a preview of
    the response so long runs can be monitored without opening the Excel file.
    """
    total_chars = sum(len(m["content"]) for m in messages)
    print(f"  [GPT] Calling {AZURE_OPENAI_DEPLOYMENT} — "
          f"{total_chars:,} total chars, max_completion_tokens=16000")
    try:
        resp = client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT,
            messages=messages,
            max_completion_tokens=16000,
        )
        TOKEN_USAGE.add_from_response(resp)
        raw = (resp.choices[0].message.content or "").strip()
        if not raw:
            print(f"  [GPT] ERROR: Empty response")
        else:
            print(f"  [GPT] OK — {len(raw)} chars: {raw[:200]!r}")
        return raw
    except Exception as e:
        print(f"  [GPT] ERROR: {type(e).__name__}: {e}")
        return ""

# ---------------------------------------------------------------------------
# arXiv fetcher
# ---------------------------------------------------------------------------

PAPER_SIGNALS = re.compile(
    r"(?:abstract|introduction|related work|references|appendix|theorem|figure|table)\b",
    re.IGNORECASE)

def is_junk(text):
    """Return True if the fetched page is a navigation/stub rather than actual paper content.

    arXiv HTML renders sometimes return the site's landing page instead of the
    paper when the ID is not yet processed by ar5iv. We reject pages that are
    too short or that contain multiple arXiv site navigation phrases without
    enough academic signal words.
    """
    junk_phrases = [
        "arXivLabs", "Papers with Code", "Help | Advanced Search",
        "Subscribe to arXiv", "Privacy Policy", "Cornell University",
        "What is ScienceCast", "What is Replicate", "Hugging Face Spaces",
    ]
    junk_count = sum(1 for p in junk_phrases if p in text[:3000])
    paper_hits = len(PAPER_SIGNALS.findall(text[:5000]))
    if len(text) < 10_000:
        print(f"  [arXiv] JUNK: only {len(text):,} chars — stub or nav page")
        return True
    if junk_count >= 2 and paper_hits < 5:
        print(f"  [arXiv] JUNK: {junk_count} nav phrases, {paper_hits} paper signals")
        return True
    return False

def fetch_appendix(title):
    """Resolve a paper by title on arXiv and return its appendix text.

    Tries ar5iv (a more reliable HTML renderer) first, then falls back to the
    official arxiv.org HTML endpoint. If no explicit appendix heading is found,
    returns the last 25% of the paper as a best-effort approximation — many
    papers place supplementary material at the end without labelling it.
    Returns None if both fetch attempts fail or the page is junk.
    """
    print(f"  [arXiv] Searching: {title[:70]}")
    try:
        query = urllib.parse.quote(f'ti:"{title}"')
        resp  = requests.get(
            f"https://export.arxiv.org/api/query?search_query={query}&max_results=1",
            timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        m = re.search(r"<id>http://arxiv\.org/abs/([^<\s]+)</id>", resp.text)
        if not m:
            print(f"  [arXiv] ERROR: No paper ID. Snippet: {resp.text[:200]}")
            return None
        arxiv_id = re.sub(r"v\d+$", "", m.group(1).strip())
        print(f"  [arXiv] Resolved ID: {arxiv_id}")
    except Exception as e:
        print(f"  [arXiv] ERROR during ID lookup: {type(e).__name__}: {e}")
        return None

    for url in [f"https://ar5iv.labs.arxiv.org/html/{arxiv_id}",
                f"https://arxiv.org/html/{arxiv_id}"]:
        print(f"  [arXiv] Fetching: {url}")
        try:
            resp = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
            print(f"  [arXiv] HTTP {resp.status_code}, raw HTML: {len(resp.text):,} chars")
            if resp.status_code != 200:
                continue

            text = re.sub(r"<script[^>]*>.*?</script>", " ", resp.text, flags=re.DOTALL)
            text = re.sub(r"<style[^>]*>.*?</style>",   " ", text,      flags=re.DOTALL)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s{3,}", "\n\n", text).strip()
            print(f"  [arXiv] After tag stripping: {len(text):,} chars")

            if is_junk(text):
                print(f"  [arXiv] Skipping {url} — junk content")
                continue

            APPENDIX_RE = re.compile(
                r"(?:^|\n\n)((?:Appendix|Supplementary\s+(?:Material|Notes?)|Appendices)\b)",
                re.IGNORECASE)
            m2 = APPENDIX_RE.search(text)
            if m2:
                appendix = text[m2.start():]
                print(f"  [arXiv] Appendix heading at char {m2.start():,} — "
                      f"full appendix: {len(appendix):,} chars")
            else:
                appendix = text[len(text) * 3 // 4:]
                print(f"  [arXiv] WARNING: No appendix heading — "
                      f"using last 25%: {len(appendix):,} chars")

            print(f"  [arXiv] Preview: {appendix[:300]!r}")
            return appendix

        except Exception as e:
            print(f"  [arXiv] ERROR at {url}: {type(e).__name__}: {e}")

    print(f"  [arXiv] ERROR: All fetch attempts failed")
    return None

# ---------------------------------------------------------------------------
# Parse GPT response
# ---------------------------------------------------------------------------

def parse(raw):
    """Parse the LLM's JSON response and apply post-processing overrides.

    Two override layers correct the most common false positives:
    - Heuristic phrase list: certain phrases (e.g. "auto interpretation",
      "dataset generation") signal that the appendix describes an *auxiliary*
      LLM prompt rather than preprocessing pseudocode, so we downgrade to MISSING.
    - Model self-report: if the LLM sets `is_primary_input_to_target_llm` in its
      JSON, that field takes precedence over both the classification and the
      heuristic, since it directly answers what we care about.
    """
    result = {
        "classification": "MISSING",
        "confidence": 0,
        "appendix_quality": "unknown",
        "key_quotes": [],
        "matched_criteria": [],
        "is_auxiliary_prompt": False,
        "diagnostic_notes": "N/A",
        "final_reason": "N/A",
        "evidence": "N/A"
    }
    if not raw:
        result["diagnostic_notes"] = "EMPTY GPT RESPONSE"
        return result

    try:
        data = json.loads(raw)
        result.update(data)
        result["evidence"] = "; ".join(data.get("key_quotes", [])) or data.get("diagnostic_notes", "N/A")
    except json.JSONDecodeError:
        result["diagnostic_notes"] = f"INVALID JSON — raw started: {raw[:100]}"

    # STRONGER HEURISTIC (kills the exact false positives we saw)
    if result["classification"] == "SEPARATE_APPENDIX":
        evidence_lower = (result.get("evidence", "") + result.get("diagnostic_notes", "")).lower()
        bad_phrases = ["auto interpretation", "monosemanticity", "question generation", "feature labeling",
                       "dataset generation", "corpus creation", "auditing", "auto-interpret"]
        if any(phrase in evidence_lower for phrase in bad_phrases):
            result["classification"] = "MISSING"
            result["is_auxiliary_prompt"] = True
            result["diagnostic_notes"] += " [HEURISTIC OVERRIDE: auxiliary prompt detected]"

    # <<<=== ADD THE OVERRIDE RIGHT HERE ===>>>
    # FINAL OVERRIDE — model’s self-reported field decides
    if result.get("is_primary_input_to_target_llm") is True:
        result["classification"] = "SEPARATE_APPENDIX"
        result["diagnostic_notes"] += " [OVERRIDE: primary pseudocode confirmed]"
    elif result.get("is_primary_input_to_target_llm") is False and result.get("is_auxiliary_prompt") is True:
        result["classification"] = "MISSING"
        result["diagnostic_notes"] += " [OVERRIDE: auxiliary only]"

    return result

# ---------------------------------------------------------------------------
# Per-paper pipeline
# ---------------------------------------------------------------------------

def process(entry):
    """Run the full classification pipeline for a single paper.

    Primary path: fetch the appendix from arXiv and classify it with the main prompt.
    Fallback path: when arXiv is unreachable or returns junk, use a knowledge-only
    prompt so the LLM answers from its training data — less reliable but better than
    skipping the paper entirely.
    """
    title    = entry["title"]
    appendix = fetch_appendix(title)

    if appendix:
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": USER_PROMPT.format(
                title=title, appendix_text=appendix)},
        ]
        source = "ARXIV"
    else:
        print(f"  [PIPELINE] No appendix retrieved — falling back to GPT training knowledge")
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": FALLBACK_PROMPT.format(title=title)},
        ]
        source = "KNOWLEDGE"

    raw    = call_gpt(messages)
    result = parse(raw)
    result["source"]            = source
    result["title"]             = title
    result["semanticscholarid"] = entry.get("semanticscholarid", "")
    return result

# ---------------------------------------------------------------------------
# Accuracy
# ---------------------------------------------------------------------------

def accuracy_report(results):
    """Compare predictions against ground_truth labels in PAPERS and return metrics.

    Only papers that have a ground_truth entry are included in the counts.
    Papers without a ground_truth label are ignored, not penalised.
    """
    gt_map = {p["title"]: p["ground_truth"] for p in PAPERS if "ground_truth" in p}
    correct, total, wrong, per_label = 0, 0, [], {}
    for r in results:
        gt = gt_map.get(r["title"])
        if not gt:
            continue
        total += 1
        per_label.setdefault(gt, {"correct": 0, "total": 0})
        per_label[gt]["total"] += 1
        if r["classification"] == gt:
            correct += 1
            per_label[gt]["correct"] += 1
        else:
            wrong.append({"title": r["title"],
                          "pred":  r["classification"],
                          "gt":    gt})
    return {"accuracy":  correct / total if total else 0,
            "correct":   correct,
            "total":     total,
            "per_label": per_label,
            "wrong":     wrong}

# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def save_excel(results, path, acc):
    """Write per-paper predictions and a Summary sheet to an Excel workbook.

    The Results sheet colour-codes each row by predicted label and marks
    correct/wrong predictions against ground truth. The Summary sheet shows
    per-label accuracy and lists all wrong predictions for quick review.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    COLOR = {
        "SEPARATE_APPENDIX": "FFEB9C",
        "MISSING":           "FCE4D6",
    }
    LABEL = {
        "SEPARATE_APPENDIX": "SEPARATE APPENDIX",
        "MISSING":           "MISSING",
    }
    gt_map    = {p["title"]: p.get("ground_truth", "") for p in PAPERS}
    gt_reason = {p["title"]: p.get("gt_reason",    "") for p in PAPERS}

    wb = Workbook()
    ws = wb.active
    ws.title = "Results"

    headers = ["#", "Title", "Prediction", "Ground Truth", "Correct?",
               "Source", "Evidence", "Reasoning", "GT Reason"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        c.fill      = PatternFill("solid", start_color="2F4F6F")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 25

    for i, r in enumerate(results, 1):
        row = i + 1
        clf = r.get("classification", "MISSING")
        gt  = gt_map.get(r["title"], "")
        ok  = (clf == gt) if gt else None
        bg  = "F9F9F9" if i % 2 == 0 else "FFFFFF"

        def cell(col, val, bold=False, fg="000000", bg_ov=None):
            c = ws.cell(row=row, column=col, value=val)
            c.font      = Font(name="Arial", size=10, bold=bold, color=fg)
            c.alignment = Alignment(
                horizontal="center" if col in (1,3,4,5,6) else "left",
                vertical="top", wrap_text=True)
            c.fill = PatternFill("solid", start_color=bg_ov or bg)

        cell(1, i)
        cell(2, r["title"])
        for col_i, key in [(3, clf), (4, gt)]:
            c = ws.cell(row=row, column=col_i, value=LABEL.get(key, key))
            c.font      = Font(bold=True, name="Arial", size=10)
            c.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
            c.fill      = PatternFill("solid", start_color=COLOR.get(key, "F2F2F2"))
        if ok is True:
            cell(5, "✓", bold=True, fg="276221", bg_ov="C6EFCE")
        elif ok is False:
            cell(5, "✗", bold=True, fg="9C0006", bg_ov="FFC7CE")
        else:
            cell(5, "—")
        src_color = {"ARXIV": "BDD7EE", "KNOWLEDGE": "FFE4B5"}.get(r.get("source",""), bg)
        c6 = ws.cell(row=row, column=6, value=r.get("source",""))
        c6.font      = Font(name="Arial", size=10, bold=True)
        c6.alignment = Alignment(horizontal="center", vertical="top")
        c6.fill      = PatternFill("solid", start_color=src_color)
        cell(7, r.get("evidence",  ""))
        cell(8, r.get("reasoning", ""))
        cell(9, gt_reason.get(r["title"], ""))
        ws.row_dimensions[row].height = 60

    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Overall Accuracy"
    ws2["B1"] = f"{acc['accuracy']:.1%}  ({acc['correct']}/{acc['total']})"
    ws2["A1"].font = ws2["B1"].font = Font(bold=True, name="Arial", size=12)
    for col, h in enumerate(["Label","Correct","Total","Accuracy"], 1):
        ws2.cell(row=3, column=col, value=h).font = Font(bold=True, name="Arial")
    r2 = 4
    for lbl, s in sorted(acc["per_label"].items()):
        a = s["correct"] / s["total"] if s["total"] else 0
        ws2.cell(row=r2, column=1, value=lbl)
        ws2.cell(row=r2, column=2, value=s["correct"])
        ws2.cell(row=r2, column=3, value=s["total"])
        ws2.cell(row=r2, column=4, value=f"{a:.1%}")
        r2 += 1
    r2 += 1
    ws2.cell(row=r2, column=1, value="Wrong predictions").font = Font(bold=True)
    r2 += 1
    for w in acc["wrong"]:
        ws2.cell(row=r2, column=1, value=w["title"])
        ws2.cell(row=r2, column=2, value=w["pred"])
        ws2.cell(row=r2, column=3, value=w["gt"])
        r2 += 1
    for col, w in [("A",10),("B",22),("C",22),("D",10)]:
        ws2.column_dimensions[col].width = w
    for col, w in [("A",5),("B",48),("C",22),("D",22),
                   ("E",10),("F",12),("G",50),("H",50),("I",50)]:
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(path)
    print(f"\n  Saved: {path}")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    results = []
    for entry in PAPERS:
        title = entry.get("title", "")
        print(f"\n{'='*60}")
        print(f"  {title[:70]}")
        print(f"{'='*60}")
        try:
            result = process(entry)
        except Exception as e:
            print(f"  [FATAL] {type(e).__name__}: {e}")
            result = {
                "title":             title,
                "semanticscholarid": entry.get("semanticscholarid", ""),
                "classification":    "MISSING",
                "evidence":          f"ERROR: {e}",
                "reasoning":         "Exception during processing",
                "source":            "ERROR",
            }
        results.append(result)
        print(f"  -> FINAL: [{result['source']}] {result['classification']}")
        time.sleep(PAUSE)

    print(f"\n{'='*60}\nSUMMARY\n{'='*60}")
    print(f"  Labels:  {dict(Counter(r['classification'] for r in results))}")
    print(f"  Sources: {dict(Counter(r['source']         for r in results))}")
    acc = accuracy_report(results)
    print(f"\n  Accuracy: {acc['accuracy']:.1%} ({acc['correct']}/{acc['total']})")
    for lbl, s in sorted(acc["per_label"].items()):
        a = s["correct"] / s["total"] if s["total"] else 0
        print(f"    {lbl:25s}: {a:.0%} ({s['correct']}/{s['total']})")
    if acc["wrong"]:
        print(f"\n  Wrong ({len(acc['wrong'])}):")
        for w in acc["wrong"]:
            print(f"    PRED={w['pred']:25s}  GT={w['gt']:25s}  {w['title'][:50]}")
    print_token_usage_report(TOKEN_USAGE, AZURE_OPENAI_DEPLOYMENT)
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results_513.xlsx")
    save_excel(results, out, acc)