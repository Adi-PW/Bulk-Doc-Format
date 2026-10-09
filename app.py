"""
QBG Doc Formatter — Streamlit front end.

Run locally:   streamlit run app.py
"""
from __future__ import annotations

import base64
import html
import os
import re

import pandas as pd
import streamlit as st

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_ROOT = "1gzMrWz-I7yrqW6S_viabr5QI4G9QY8QL"
CREATOR_EMAIL = "aditya.raj3@pw.live"

st.set_page_config(page_title="QBG Doc Formatter", page_icon="⚖️", layout="wide",
                   initial_sidebar_state="collapsed")

# The code works in two layouts:
#   package: qbg/*.py and assets/* next to app.py   (as in the zip)
#   flat:    every file directly next to app.py     (GitHub browser upload without folders)
CODE_FILES = ["docx_reader.py", "parser.py", "builder.py", "google_services.py", "pipeline.py"]
ASSET_FILES = ["question_template.docx", "creator.png"]
PACKAGE_LAYOUT = all(os.path.exists(os.path.join(APP_DIR, "qbg", f)) for f in CODE_FILES)


def asset_path(name: str) -> str:
    for cand in (os.path.join(APP_DIR, "assets", name), os.path.join(APP_DIR, name)):
        if os.path.exists(cand):
            return cand
    return ""


_missing = [] if PACKAGE_LAYOUT else [f for f in CODE_FILES
                                      if not os.path.exists(os.path.join(APP_DIR, f))]
_missing += [f for f in ASSET_FILES if not asset_path(f)]
if _missing:
    st.error("**Some project files are missing from the GitHub repository**, so the app can't start.\n\n"
             "Missing: " + ", ".join(f"`{f}`" for f in _missing) + "\n\n"
             "Upload them next to `app.py` (GitHub: **Add file → Upload files**). "
             "Streamlit redeploys automatically.")
    st.stop()

if PACKAGE_LAYOUT:
    from qbg.google_services import GoogleDrive, build_credentials  # noqa: E402
    from qbg.parser import parse_docx, summarize  # noqa: E402
    from qbg.pipeline import FileTask, Job, count_rows, to_csv, to_excel  # noqa: E402
else:
    from google_services import GoogleDrive, build_credentials  # noqa: E402
    from parser import parse_docx, summarize  # noqa: E402
    from pipeline import FileTask, Job, count_rows, to_csv, to_excel  # noqa: E402


# --------------------------------------------------------------------------- #
# Config / shared state
# --------------------------------------------------------------------------- #
def secret(key: str, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


ROOT_FOLDER_ID = secret("ROOT_FOLDER_ID", DEFAULT_ROOT)
DEFAULT_WORKERS = int(secret("MAX_WORKERS", 12))


@st.cache_resource
def job_registry() -> dict:
    """Process-wide registry so a page refresh can re-attach to a running job."""
    return {}


def workspace_factory():
    creds = build_credentials(st.secrets)
    return GoogleDrive(creds, ROOT_FOLDER_ID)


def credentials_configured() -> bool:
    try:
        return "google_oauth" in st.secrets or "gcp_service_account" in st.secrets
    except Exception:
        return False


@st.cache_data(show_spinner=False)
def avatar_b64() -> str:
    with open(asset_path("creator.png"), "rb") as fh:
        return base64.b64encode(fh.read()).decode()


# --------------------------------------------------------------------------- #
# Styling
# --------------------------------------------------------------------------- #
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,500;12..96,700;12..96,800&family=Figtree:wght@400;500;600&display=swap');
:root {
  --ink:#1F2A5C; --ink-2:#2F3E80; --paper:#F4F6FB; --card:#FFFFFF;
  --seal:#C8323C; --ledger:#1E7B4F; --slate:#5B6478; --line:#DCE1EC;
}
html, body, [class*="css"], .stMarkdown, .stText, label, input, textarea, button {
  font-family: 'Figtree', system-ui, -apple-system, 'Segoe UI', sans-serif !important;
}
.block-container { padding-top: 2.2rem; padding-bottom: 6rem; max-width: 1180px; }
h1, h2, h3, .qbg-display { font-family: 'Bricolage Grotesque', 'Figtree', sans-serif !important; }

/* ---------- hero ---------- */
.qbg-hero { display:grid; grid-template-columns: 1.5fr 1fr; gap: 2rem; align-items:center;
  background: var(--ink); color:#EEF1FA; border-radius: 22px; padding: 2.4rem 2.6rem;
  position: relative; overflow: hidden; }
.qbg-hero::after { content:""; position:absolute; inset:0;
  background-image: repeating-linear-gradient(0deg, rgba(255,255,255,.035) 0 1px, transparent 1px 30px);
  pointer-events:none; }
.qbg-hero h1 { font-size: clamp(2rem, 3.6vw, 3.1rem); line-height:1.04; margin:0 0 .9rem;
  font-weight:800; letter-spacing:-0.02em; color:#FFFFFF; }
.qbg-hero p { font-size:1.06rem; line-height:1.55; color:#C9D0E8; margin:0; max-width: 38rem; }
.qbg-stampwrap { display:flex; justify-content:center; }
.qbg-stamp { border: 4px solid var(--seal); color: var(--seal); border-radius: 14px;
  padding: .9rem 1.3rem .8rem; transform: rotate(-7deg); background: rgba(255,255,255,.96);
  box-shadow: inset 0 0 0 3px #fff, inset 0 0 0 5px var(--seal);
  text-align:center; animation: qbg-press .9s cubic-bezier(.2,1.4,.4,1) .25s both; }
.qbg-stamp .s1 { font-family:'Bricolage Grotesque',sans-serif; font-weight:800;
  font-size: clamp(1.25rem, 2.3vw, 1.9rem); letter-spacing: .04em; }
.qbg-stamp .s2 { font-size:.86rem; font-weight:600; margin-top:.15rem; color:#9A2630; }
@keyframes qbg-press { 0% { transform: rotate(-7deg) scale(1.6); opacity:0; }
  60% { opacity:1; } 100% { transform: rotate(-7deg) scale(1); opacity:1; } }
@media (prefers-reduced-motion: reduce) { .qbg-stamp { animation:none; } }
@media (max-width: 760px) { .qbg-hero { grid-template-columns: 1fr; padding: 1.8rem; } }

/* ---------- steps ---------- */
.qbg-step { display:flex; gap:.8rem; align-items:baseline; margin: 2.2rem 0 .6rem; }
.qbg-step .n { font-family:'Bricolage Grotesque',sans-serif; font-weight:800; color:var(--seal);
  font-size:1.35rem; min-width:1.6rem; }
.qbg-step .t { font-family:'Bricolage Grotesque',sans-serif; font-weight:700; font-size:1.35rem;
  color:var(--ink); }
.qbg-step .h { color:var(--slate); font-size:.95rem; }

/* ---------- tallies ---------- */
.qbg-tally { display:grid; grid-template-columns: repeat(4, 1fr); gap: 0; margin:.6rem 0 1rem;
  border:1px solid var(--line); border-radius:16px; background:var(--card); overflow:hidden; }
.qbg-tally > div { padding: 1rem 1.2rem; border-right:1px solid var(--line); }
.qbg-tally > div:last-child { border-right:none; }
.qbg-tally .v { font-family:'Bricolage Grotesque',sans-serif; font-weight:800; font-size:2rem;
  color:var(--ink); line-height:1.1; }
.qbg-tally .v.ok { color: var(--ledger); } .qbg-tally .v.bad { color: var(--seal); }
.qbg-tally .k { color:var(--slate); font-size:.9rem; }
@media (max-width: 760px) { .qbg-tally { grid-template-columns: repeat(2, 1fr); } }
.qbg-phase { color: var(--ink-2); font-weight:600; margin:.2rem 0 .4rem; }

/* ---------- buttons ---------- */
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
  background: var(--ink); border-color: var(--ink); border-radius: 12px; font-weight:600; }
.stButton > button[kind="primary"]:hover { background: var(--ink-2); border-color: var(--ink-2); }
.stButton > button, .stDownloadButton > button { border-radius: 12px; }
.stButton > button[kind="primary"]:disabled { background: #C9CFE0; border-color: #C9CFE0; color: #4A5270; }
.qbg-hero { margin-bottom: 1.2rem; }
button:focus-visible { outline: 3px solid #8FA2FF !important; outline-offset: 2px; }
[data-testid="stFileUploaderDropzone"] { border-radius: 16px; border: 1.5px dashed #AEB8D3;
  background: var(--card); }

/* ---------- creator badge ---------- */
.qbg-creator { position: fixed; left: 18px; bottom: 18px; z-index: 9999;
  display:flex; align-items:center; gap:.55rem; text-decoration:none !important;
  background: var(--card); color: var(--ink) !important; border:1px solid var(--line);
  padding: .32rem .85rem .32rem .32rem; border-radius: 999px;
  box-shadow: 0 6px 22px rgba(31,42,92,.16); font-size:.88rem; font-weight:600; }
.qbg-creator img { width:30px; height:30px; border-radius:50%; object-fit:cover; }
.qbg-creator span b { color: var(--seal); font-weight:700; }
.qbg-creator:hover { border-color: var(--ink); }
"""


def inject_css():
    # Markdown ends an HTML block at a blank line, so comments/blank lines are stripped.
    lines = [l for l in CSS.splitlines() if l.strip() and not l.strip().startswith("/*")]
    st.markdown("<style>" + "\n".join(lines) + "</style>", unsafe_allow_html=True)
    st.markdown(
        f'<a class="qbg-creator" href="mailto:{CREATOR_EMAIL}" title="{CREATOR_EMAIL}">'
        f'<img src="data:image/png;base64,{avatar_b64()}" alt="Aditya Raj">'
        f'<span>Creator: <b>Aditya Raj</b></span></a>', unsafe_allow_html=True)


inject_css()


# --------------------------------------------------------------------------- #
# Hero
# --------------------------------------------------------------------------- #
st.markdown("""
<div class="qbg-hero">
  <div>
    <h1>Turn test-paper docs into ready-to-ingest questions</h1>
    <p>Drop in the Word files your team already writes. Each question, its options,
    the answer key and the explanation are pulled out, laid into the QBG format,
    saved to Drive as a numbered doc, with an Excel record per file.</p>
  </div>
  <div class="qbg-stampwrap">
    <div class="qbg-stamp"><div class="s1">QBGDOC000001</div><div class="s2">Ready to ingest</div></div>
  </div>
</div>
""", unsafe_allow_html=True)


def step(n: int, title: str, hint: str = ""):
    st.markdown(f'<div class="qbg-step"><span class="n">{n}</span><span class="t">{title}</span>'
                f'<span class="h">{html.escape(hint)}</span></div>', unsafe_allow_html=True)


def tally(cells):
    if not cells:
        return
    inner = "".join(f'<div><div class="v {c}">{v}</div><div class="k">{k}</div></div>'
                    for k, v, c in cells)
    st.markdown(f'<div class="qbg-tally" style="grid-template-columns: repeat({len(cells)}, 1fr)">'
                f'{inner}</div>', unsafe_allow_html=True)


def language_tally(c: dict, auto: int):
    """Second row of counts: languages + automatically written solutions."""
    cells = []
    be, bh = c.get("bilingual_english", 0), c.get("bilingual_hindi", 0)
    if be or bh:
        mismatch = "bad" if be != bh else ""
        cells += [("Bilingual · English", be, mismatch), ("Bilingual · Hindi", bh, mismatch)]
    if c.get("english_only"):
        cells.append(("English-only", c["english_only"], ""))
    if c.get("hindi_only"):
        cells.append(("Hindi-only", c["hindi_only"], ""))
    cells.append(("Solutions auto-written", auto, ""))
    tally(cells)
    if be != bh:
        st.warning(f"English and Hindi versions differ by {abs(be - bh)}: some questions have only one "
                   "language. They are marked Failure with 'version not found' in Remarks.")


def safe_name(s: str) -> str:
    s = re.sub(r'[\\/:*?"<>|\n\r\t]+', " ", str(s or "")).strip()
    return re.sub(r"\s{2,}", " ", s)


# --------------------------------------------------------------------------- #
# Re-attach to a running job after refresh (?run=ID)
# --------------------------------------------------------------------------- #
registry = job_registry()
run_param = st.query_params.get("run")
if "job_id" not in st.session_state and run_param and run_param in registry:
    st.session_state.job_id = run_param


def current_job() -> Job | None:
    jid = st.session_state.get("job_id")
    return registry.get(jid) if jid else None


job = current_job()
locked = job is not None and job.active

if not credentials_configured():
    st.warning("Google credentials are not configured yet, so files can be checked but not "
               "uploaded. Add them under app Settings → Secrets (see README).")

# --------------------------------------------------------------------------- #
# Step 1: files
# --------------------------------------------------------------------------- #
step(1, "Add Word files", "  .docx or .doc, as many as you like")
uploads = st.file_uploader("Word files", type=["docx", "doc"], accept_multiple_files=True,
                           label_visibility="collapsed", disabled=locked,
                           key=f"uploader_{st.session_state.get('uploader_gen', 0)}")

# --------------------------------------------------------------------------- #
# Step 2: category + file name per file
# --------------------------------------------------------------------------- #
edited = None
if uploads:
    step(2, "Name each file", "  category decides the Drive folder; file name becomes the sub-folder")
    c1, c2 = st.columns([2, 1])
    with c1:
        default_cat = st.text_input("QBG Category for all files", value="Judiciary", disabled=locked,
                                    help="Pre-fills every row below. You can still change any row.")
    with c2:
        st.write("")
        st.write("")
        apply_all = st.button("Apply to every row", disabled=locked, width="stretch")

    sig = tuple((u.name, u.size) for u in uploads)
    if st.session_state.get("table_sig") != sig or apply_all:
        prev = st.session_state.get("last_edited", st.session_state.get("table_df"))
        rows = []
        for u in uploads:
            stem = os.path.splitext(u.name)[0]
            keep = None
            if prev is not None and not apply_all:
                m = prev[prev["Uploaded file"] == u.name]
                keep = m.iloc[0] if len(m) else None
            rows.append({
                "Uploaded file": u.name,
                "QBG Category": keep["QBG Category"] if keep is not None else default_cat,
                "File Name": keep["File Name"] if keep is not None else stem,
            })
        st.session_state.table_df = pd.DataFrame(rows)
        st.session_state.table_sig = sig
        st.session_state.editor_gen = st.session_state.get("editor_gen", 0) + 1

    edited = st.data_editor(
        st.session_state.table_df, hide_index=True, width="stretch", disabled=locked,
        key=f"editor_{st.session_state.get('editor_gen', 0)}",
        column_config={
            "Uploaded file": st.column_config.TextColumn(disabled=True, width="medium"),
            "QBG Category": st.column_config.TextColumn(required=True, width="medium"),
            "File Name": st.column_config.TextColumn(required=True, width="large"),
        })
    st.session_state.last_edited = edited

    with st.expander("Output settings"):
        s1, s2, s3 = st.columns(3)
        language = s1.text_input("Input Language", value="English", disabled=locked)
        qtype = s2.text_input("Question Type", value="SCQ", disabled=locked)
        workers = s3.slider("Parallel uploads", 2, 24, DEFAULT_WORKERS, disabled=locked,
                            help="Higher is faster. If Google briefly throttles, the app "
                                 "waits and retries automatically.")
else:
    language, qtype, workers = "English", "SCQ", DEFAULT_WORKERS


def validate(df: pd.DataFrame) -> list[str]:
    problems = []
    for _, r in df.iterrows():
        if not safe_name(r["QBG Category"]):
            problems.append(f"{r['Uploaded file']}: QBG Category is empty")
        if not safe_name(r["File Name"]):
            problems.append(f"{r['Uploaded file']}: File Name is empty")
    dup = df.assign(_k=df["QBG Category"].map(safe_name).str.lower() + "/" +
                    df["File Name"].map(safe_name).str.lower())
    for k, n in dup["_k"].value_counts().items():
        if n > 1:
            problems.append(f"{n} files share the same Category and File Name ({k}); "
                            "give each a distinct File Name so their questions land in separate folders")
    return problems


# --------------------------------------------------------------------------- #
# Step 3: check + run
# --------------------------------------------------------------------------- #
if uploads and edited is not None:
    step(3, "Check, then format", "  checking is instant and touches nothing in Drive")
    problems = validate(edited)
    for p in problems:
        st.error(p)

    b1, b2, _ = st.columns([1, 1.3, 2])
    check = b1.button("Check parsing", disabled=locked or bool(problems), width="stretch")
    start = b2.button("Format and upload", type="primary", width="stretch",
                      disabled=locked or bool(problems) or not credentials_configured())

    if check:
        by_name = {u.name: u for u in uploads}
        total = ready = 0
        report = []
        all_counts: dict = {}
        for _, r in edited.iterrows():
            u = by_name[r["Uploaded file"]]
            if u.name.lower().endswith(".doc"):
                report.append((r["File Name"], None, None, "Legacy .doc: converted through Google Drive "
                               "during the run, so it cannot be checked offline", [], {}))
                continue
            try:
                qs, keys = parse_docx(u.getvalue())
            except Exception as e:
                report.append((r["File Name"], 0, 0, f"Could not read file: {e}", [], {}))
                continue
            ok = sum(q.ok for q in qs)
            total += len(qs)
            ready += ok
            fc = summarize(qs)
            for key, val in fc.items():
                all_counts[key] = all_counts.get(key, 0) + val
            note = "" if keys == len(qs) else f"Answer keys in file: {keys}, questions detected: {len(qs)}"
            issues = [q for q in qs if not q.ok or q.warnings]
            report.append((r["File Name"], len(qs), ok, note, issues, fc))
        tally([("Questions found", total, ""), ("Ready to format", ready, "ok"),
               ("Need attention", total - ready, "bad" if total - ready else ""), ("Files", len(report), "")])
        language_tally(all_counts, all_counts.get("auto_solutions", 0))
        for name, n, ok, note, issues, fc in report:
            label = f"{name}  —  {n} found, {ok} ready" if n is not None else f"{name}  —  {note}"
            with st.expander(label, expanded=bool(issues) and len(report) == 1):
                if n is not None and fc:
                    st.caption(" · ".join(f"{k}: {v}" for k, v in count_rows(fc).items())
                               + f" · Solutions to auto-write: {fc.get('auto_solutions', 0)}")
                if note and n is not None:
                    st.warning(note)
                if not issues:
                    st.success("Every question parsed cleanly.")
                for q in issues:
                    kind = st.error if not q.ok else st.info
                    kind(f"**{q.label}** — {q.remarks().split(': ', 1)[-1]}\n\n"
                         f"> {q.body_text[:220].replace(chr(10), ' ')}")

    if start:
        by_name = {u.name: u for u in uploads}
        tasks = [FileTask(index=i, uploaded_name=r["Uploaded file"],
                          data=by_name[r["Uploaded file"]].getvalue(),
                          category=safe_name(r["QBG Category"]), file_name=safe_name(r["File Name"]))
                 for i, (_, r) in enumerate(edited.iterrows())]
        new_job = Job(tasks, workspace_factory, language=language or "English",
                      qtype=qtype or "SCQ", max_workers=int(workers))
        registry[new_job.id] = new_job
        st.session_state.job_id = new_job.id
        st.query_params["run"] = new_job.id
        new_job.start()
        st.rerun()


# --------------------------------------------------------------------------- #
# Progress + results
# --------------------------------------------------------------------------- #
def results_frame(j: Job) -> pd.DataFrame:
    rows = j.results()
    return pd.DataFrame([r.as_list() for r in rows], columns=[
        "QBG Category", "File Name", "Question Serial No.", "Original Doc Link",
        "Question Doc Link", "Status", "Remarks"])


def render_job(j: Job):
    step(4, "Progress" if j.active else "Results", f"  {j.id}")
    total = j.total_detected
    pct = (j.processed / total) if total else (0.0 if j.active else 1.0)
    st.markdown(f'<div class="qbg-phase">{html.escape(j.phase)}</div>', unsafe_allow_html=True)
    st.progress(min(max(pct, 0.0), 1.0))
    mins, secs = divmod(int(j.elapsed), 60)
    tally([("Questions given", total, ""), ("Formatted", j.formatted, "ok"),
           ("Failed", j.failed, "bad" if j.failed else ""), ("Elapsed", f"{mins}:{secs:02d}", "")])
    if j.counts:
        language_tally(j.counts, j.auto_written)
    if j.status == "error":
        st.error(f"The run stopped because of an error: {j.error}. Questions completed before the "
                 "error are listed below and were recorded in Drive where possible.")

    df = results_frame(j)
    if not df.empty:
        st.dataframe(df, hide_index=True, width="stretch", height=360, column_config={
            "Original Doc Link": st.column_config.LinkColumn(display_text="Open original"),
            "Question Doc Link": st.column_config.LinkColumn(display_text="Open doc"),
        })
    excel_links = [(fs.task.file_name, fs.excel_link) for fs in j.files if fs.excel_link]
    if excel_links:
        st.markdown("**Excel records saved in Drive:** " + " · ".join(
            f"[{html.escape(n)}]({u})" for n, u in excel_links))
    with st.expander("Activity log"):
        st.code("\n".join(j.log[-80:]) or "Starting…", language=None)


@st.fragment(run_every=1.0)
def live_panel():
    j = current_job()
    if j is None:
        return
    if not j.active:
        st.rerun()                      # hand over to the static results view
    render_job(j)
    if j.status == "stopping":
        st.button("Stopping…", disabled=True)
    elif st.button("Stop processing", type="secondary", help="Uploads already in flight finish; "
                   "everything else is marked 'Not processed' so you can re-run it."):
        j.stop()


job = current_job()
if job is not None:
    if job.active:
        live_panel()
    else:
        render_job(job)
        rows = job.results()
        summary = {
            "Run ID": job.id, "Status": job.status,
            "Files": len(job.files), "Total questions given": job.total_detected,
            "Total questions formatted": job.formatted, "Failed / not processed": job.failed,
            "Answer keys found in files": job.total_answer_keys,
            "Solutions written automatically": job.auto_written,
            **count_rows(job.counts),
            "Processing time (s)": round(job.elapsed, 1),
        }
        for fs in job.files:
            if fs.excel_link:
                summary[f"Excel record: {fs.task.file_name}"] = fs.excel_link
        st.markdown(f"**{job.formatted} of {job.total_detected}** questions formatted "
                    f"across {len(job.files)} file(s).")
        d1, d2, d3 = st.columns([1, 1, 1.2])
        d1.download_button("Download Excel", to_excel(rows, summary), file_name=f"{job.id}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch")
        d2.download_button("Download CSV", to_csv(rows), file_name=f"{job.id}.csv", mime="text/csv",
                           width="stretch")
        if d3.button("Start a new batch", width="stretch"):
            st.session_state.pop("job_id", None)
            st.session_state.pop("table_df", None)
            st.session_state.pop("table_sig", None)
            st.session_state.pop("last_edited", None)
            st.session_state.uploader_gen = st.session_state.get("uploader_gen", 0) + 1
            st.query_params.clear()
            st.rerun()
