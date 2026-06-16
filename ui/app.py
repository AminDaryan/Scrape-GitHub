"""Streamlit UI for the paper-reproducibility checkers.

Launch from the repo root (with the venv active):

    streamlit run ui/app.py

Three tabs — 5.1 (code availability), 5.2 (usability/popularity), and the
Beall's predatory-venue check.  Upload a papers JSON (a list, or a single paper
object), pick a check, run it, preview the result, and download the Excel.

Each check runs the existing checker script as a subprocess with the upload
injected via an env var (see ui/runners.py), so the UI and the CLI always
produce identical output.
"""

import json
import os

import openpyxl
import streamlit as st
from dotenv import load_dotenv

import runners

load_dotenv(runners.REPO_ROOT / ".env")

st.set_page_config(page_title="Paper reproducibility checks", page_icon="📄", layout="wide")

# ── Example input shown to the user per section ─────────────────────────────
_GH_EXAMPLE = json.dumps(
    [{"title": "Example paper", "repo": "https://github.com/pallets/flask",
      "semanticscholarid": ""}], indent=2)
_S2_EXAMPLE = json.dumps(
    [{"title": "Example paper", "venue": "Journal of X", "year": 2023,
      "externalIds": {"DOI": "10.1234/x"},
      "publicationVenue": {"name": "Journal of X", "type": "journal",
                           "url": "https://www.scirp.org/journal/x",
                           "issn": "1234-5678", "alternate_urls": []},
      "openAccessPdf": {"url": ""}}], indent=2)


def env_sidebar():
    st.sidebar.header("Environment")
    token = bool(os.getenv("GITHUB_TOKEN"))
    st.sidebar.write(("✅" if token else "⚠️") + " `GITHUB_TOKEN` "
                     + ("set" if token else "missing (60 req/hr limit)"))
    st.sidebar.write(f"**LLM provider:** `{os.getenv('LLM_PROVIDER', 'azure')}`")
    st.sidebar.write(f"**Deployment:** `{os.getenv('AZURE_OPENAI_DEPLOYMENT', '(unset)')}`")
    st.sidebar.caption("Keys are read from the repo's `.env` by each checker. "
                       "LLM checks cost tokens; large lists can be slow.")


def get_papers(section_key: str, corpus: bool):
    """Render the upload/paste widget and return a parsed list of papers, or None."""
    example = _S2_EXAMPLE if corpus else _GH_EXAMPLE
    fmt = ("Semantic Scholar **paper records** (with `publicationVenue`)"
           if corpus else "papers with **GitHub links** (`title` + `repo`)")
    st.caption(f"Input format: {fmt}. Upload a JSON **list**, or a **single** paper object.")
    with st.expander("See expected format"):
        st.code(example, language="json")
    mode = st.radio("Provide papers via", ["Upload file", "Paste JSON"],
                    horizontal=True, key=f"mode_{section_key}")
    raw = None
    if mode == "Upload file":
        up = st.file_uploader("papers.json", type=["json"], key=f"file_{section_key}")
        if up is not None:
            raw = up.getvalue().decode("utf-8")
    else:
        raw = st.text_area("Paste JSON here", height=160, key=f"paste_{section_key}",
                           placeholder=example)

    if not raw or not raw.strip():
        return None
    try:
        items = runners.normalize_upload(json.loads(raw))
    except json.JSONDecodeError as e:
        st.error(f"Invalid JSON: {e}")
        return None
    except ValueError as e:
        st.error(str(e))
        return None
    st.success(f"Loaded {len(items)} paper(s).")
    return items


def venue_list_input(label: str, key: str):
    """Optional JSON list of venues for whitelist/blacklist. Returns list or None."""
    raw = st.text_area(label, height=90, key=key,
                       placeholder='[{"name": "IEEE", "domain": "ieee.org"}, "Nature"]')
    if not raw or not raw.strip():
        return None
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        st.error(f"{label}: invalid JSON ({e})")
        return None
    if isinstance(parsed, dict):
        parsed = [parsed]
    if not isinstance(parsed, list):
        st.error(f"{label}: must be a JSON list of venues.")
        return None
    return parsed


def render_result(res: runners.RunResult, checker: runners.Checker):
    if res.ok:
        st.success(f"Done — produced `{res.output_path.name}`.")
        st.download_button("⬇️ Download Excel", data=res.output_bytes,
                           file_name=res.output_path.name, key=f"dl_{checker.id}",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        _preview(res.output_bytes)
    else:
        st.error(f"The checker exited with code {res.returncode} and no output file was found.")
    with st.expander("Run log (stdout / stderr)", expanded=not res.ok):
        if res.stdout:
            st.text(res.stdout[-6000:])
        if res.stderr:
            st.caption("stderr")
            st.text(res.stderr[-4000:])


def _preview(xlsx_bytes: bytes, max_rows: int = 40):
    """Show sheet names + the first sheet's first rows as a table."""
    import io
    try:
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    except Exception as e:                      # never let a preview crash the page
        st.caption(f"(Preview unavailable: {e})")
        return
    st.caption("Sheets: " + ", ".join(wb.sheetnames))
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return
    headers = [str(h) if h is not None else f"col{i}" for i, h in enumerate(rows[0])]
    data = [dict(zip(headers, [("" if v is None else v) for v in r])) for r in rows[1:max_rows + 1]]
    st.dataframe(data, use_container_width=True)
    if len(rows) > max_rows + 1:
        st.caption(f"(showing first {max_rows} of {len(rows) - 1} rows — download for the full sheet)")


def section_tab(section: str):
    checkers = runners.checkers_for(section)
    labels = [c.label for c in checkers]
    label = st.selectbox("Check to run", labels, key=f"sel_{section}")
    checker = checkers[labels.index(label)]

    badges = []
    if checker.needs_llm:
        badges.append("needs LLM keys (costs tokens)")
    if not checker.needs_llm:
        badges.append("no LLM — GitHub API only")
    st.caption(" · ".join(badges))
    if checker.note:
        st.info(checker.note)

    items = get_papers(f"{section}_{checker.id}", corpus=checker.corpus_based)

    aux_lists = None
    extra_args = None
    env_overrides = None
    if section in ("5.1", "5.2"):
        st.caption("Every run validates the input first (a data-quality report appears in the "
                   "run log): missing/duplicate/non-GitHub repo links, etc.")
        if st.checkbox("Also check each repo link is live (HEAD request; slower)",
                       key=f"live_{checker.id}"):
            env_overrides = {"CHECK_REPO_LIVENESS": "1"}
    if section == "bealls":
        with st.expander("Whitelist / blacklist (optional)"):
            st.caption(
                "**Whitelist** — only papers whose venue is in this list are checked; "
                "the rest are marked *out_of_scope*. Empty = check everything.  "
                "**Blacklist** — these venues are added to Beall's List as predatory.  "
                'Format: a JSON list of `{"name": …, "domain": …}` (a plain string works as a name).')
            wl = venue_list_input("Whitelist venues (JSON)", f"wl_{checker.id}")
            bl = venue_list_input("Blacklist venues (JSON)", f"bl_{checker.id}")
        aux_lists = {"--whitelist": wl, "--blacklist": bl}
        if st.checkbox("Cross-check DOIs against Crossref (catches wrong S2 venue/ISSN; slower)",
                       key=f"cr_{checker.id}"):
            extra_args = ["--crossref", "all"]
        st.caption("Offline data-quality checks always run (a 'Data quality' column + sheet); "
                   "Crossref adds an authoritative DOI cross-check.")

    run = st.button("▶ Run check", type="primary", key=f"run_{checker.id}",
                    disabled=items is None)
    if run and items is not None:
        with st.spinner(f"Running {checker.label} on {len(items)} paper(s)… "
                        "this can take a while for large lists or LLM checks."):
            try:
                res = runners.run_checker(checker, items, aux_lists=aux_lists,
                                          extra_args=extra_args, env_overrides=env_overrides)
            except Exception as e:                 # surface any launch failure
                st.session_state[f"res_{checker.id}"] = None
                st.exception(e)
                return
        st.session_state[f"res_{checker.id}"] = res

    res = st.session_state.get(f"res_{checker.id}")
    if res is not None:
        render_result(res, checker)


st.title("📄 Paper reproducibility & venue checks")
st.caption("Upload a list of papers (or one paper) and run any check. Output matches the CLI scripts exactly.")
env_sidebar()

tab51, tab52, tab_b = st.tabs(["5.1 — Code availability", "5.2 — Usability & popularity",
                               "Beall's predatory-venue"])
with tab51:
    st.subheader("Code availability & documentation")
    section_tab("5.1")
with tab52:
    st.subheader("Practitioner usability & popularity")
    section_tab("5.2")
with tab_b:
    st.subheader("Beall's List predatory-venue check")
    st.caption("Optionally scope with a whitelist or extend the list with a blacklist "
               "(see the expander after choosing a check).")
    section_tab("bealls")
