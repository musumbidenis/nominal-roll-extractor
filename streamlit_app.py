#!/usr/bin/env python3
"""
streamlit_app.py — Web UI for the TVET CDACC Nominal Roll → Marksheet tool.

Run locally:
    streamlit run streamlit_app.py

Requires: streamlit, pdfplumber, openpyxl
"""

import io
import os
import re
import tempfile
import zipfile
from collections import Counter
from datetime import datetime

import streamlit as st

from extract_nominal import extract
from marksheet_excel import build_marksheet_per_unit
from registration_excel import build_class_forms
from summative_excel import build_summative_per_unit

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Nominal Roll → Marksheet",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="collapsed",
)

_VERSION = "v1.0"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Hide Streamlit chrome ──────────────────────────────────────────────── */
#MainMenu, footer,
[data-testid="stDeployButton"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stAppCreatorBadge"],
.viewerBadge_container__r5tak,
.viewerBadge_link__qRIco,
iframe[title="st_app_creator_badge"] { display: none !important; }
a[href*="streamlit.io"]              { display: none !important; }

/* ── Layout ─────────────────────────────────────────────────────────────── */
.block-container {
    padding-top: 0 !important;
    padding-bottom: 2rem;
    max-width: 1150px;
    overflow-x: hidden;
}

/* ── App header ─────────────────────────────────────────────────────────── */
.app-header {
    background: linear-gradient(135deg, #1a252f 0%, #2c3e50 100%);
    padding: 14px 26px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 4px;
    margin-left: -5rem;
    margin-right: -5rem;
    margin-bottom: 1.4rem;
}
.app-header h1 {
    margin: 0; padding: 0;
    font-size: 1.3rem; font-weight: 700;
    color: #fff; letter-spacing: -0.2px;
}
.app-header .ver {
    font-size: 0.7rem;
    color: rgba(255,255,255,0.45);
    align-self: flex-start;
    margin-top: 3px;
    white-space: nowrap;
}

/* ── Section label ──────────────────────────────────────────────────────── */
.sec-lbl {
    font-size: 0.67rem; font-weight: 700;
    letter-spacing: 1px; text-transform: uppercase;
    color: #6B7280; margin: 0 0 7px; padding: 0;
}

/* ── Metric cards ───────────────────────────────────────────────────────── */
.metrics-row {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin-bottom: 0.9rem;
}
.metric-card {
    flex: 1 1 0;
    min-width: 110px;
    background: #F8FAFB;
    border: 1px solid rgba(0,0,0,0.07);
    border-radius: 8px;
    padding: 13px 14px 11px;
}
.metric-card .num {
    font-size: 1.9rem; font-weight: 800;
    line-height: 1.1; margin-bottom: 5px;
}
.metric-card .lbl {
    font-size: 0.66rem; font-weight: 600;
    letter-spacing: 0.5px; text-transform: uppercase;
    color: #6B7280; line-height: 1.3;
}

/* ── Activity panel ─────────────────────────────────────────────────────── */
.act-header {
    display: flex; align-items: center;
    justify-content: space-between; margin-bottom: 8px;
}
.act-title { font-size: 0.88rem; font-weight: 700; color: #111827; }
.badge-done {
    background: #D1FAE5; color: #065F46;
    font-size: 0.67rem; font-weight: 700;
    padding: 2px 9px; border-radius: 12px; letter-spacing: 0.3px;
}
.badge-wait {
    background: #F3F4F6; color: #6B7280;
    font-size: 0.67rem; font-weight: 700;
    padding: 2px 9px; border-radius: 12px;
}
.log-box {
    font-family: Consolas, 'Courier New', monospace;
    font-size: 0.78rem;
    background: #FAFAFA;
    border: 1px solid rgba(0,0,0,0.07);
    border-radius: 6px;
    padding: 10px 13px;
    max-height: 230px;
    overflow-y: auto;
    line-height: 1.7;
    margin-top: 6px;
}

/* ── Export panel ───────────────────────────────────────────────────────── */
.export-title {
    font-size: 0.93rem; font-weight: 700;
    color: #111827; margin: 0 0 2px;
}
.export-caption {
    font-size: 0.75rem; color: #6B7280;
    margin: 0 0 10px; line-height: 1.4;
}
.fn-label {
    font-size: 0.67rem; font-weight: 600;
    letter-spacing: 0.5px; text-transform: uppercase;
    color: #6B7280; margin: 8px 0 3px;
}
.stDownloadButton > button {
    width: 100% !important;
    font-weight: 600 !important;
    border-radius: 6px !important;
}
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 10px !important; }

/* ── Dark mode ──────────────────────────────────────────────────────────── */
@media (prefers-color-scheme: dark) {
    .sec-lbl    { color: #9CA3AF; }
    .metric-card { background: #1E2530; border-color: rgba(255,255,255,0.08); }
    .metric-card .lbl { color: #9CA3AF; }
    .act-title  { color: #F3F4F6; }
    .badge-done { background: #064E3B; color: #6EE7B7; }
    .badge-wait { background: #374151; color: #9CA3AF; }
    .log-box { background: #161B22; border-color: rgba(255,255,255,0.08); }
    .export-title   { color: #F3F4F6; }
    .export-caption { color: #9CA3AF; }
    .fn-label       { color: #9CA3AF; }
}

/* ── Responsive: tablet ─────────────────────────────────────────────────── */
@media (max-width: 768px) {
    .app-header { margin-left: -2.5rem; margin-right: -2.5rem; }
    .metric-card .num { font-size: 1.6rem; }
}

/* ── Responsive: mobile ─────────────────────────────────────────────────── */
@media (max-width: 576px) {
    .app-header { margin-left: -1rem; margin-right: -1rem; padding: 10px 14px; }
    .app-header h1 { font-size: 1.05rem; }
    .metrics-row { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; }
    .metric-card { min-width: unset; padding: 10px 11px 9px; }
    .metric-card .num { font-size: 1.45rem; }
    [data-testid="stHorizontalBlock"] { flex-direction: column !important; }
    [data-testid="stColumn"] {
        width: 100% !important; min-width: 100% !important;
        flex: 1 0 100% !important;
    }
    .log-box { max-height: 160px; font-size: 0.74rem; }
    .sec-lbl { margin-bottom: 5px; }
}
</style>
""", unsafe_allow_html=True)

# ── Colour palette ────────────────────────────────────────────────────────────
_C = {
    "primary": "#2c3e50",
    "info":    "#3498db",
    "success": "#18bc9c",
    "warning": "#f39c12",
    "danger":  "#e74c3c",
}

_LOG_COLORS = {
    "info":    "#999999", "step":    "#1565C0", "found":   "#2E7D32",
    "success": "#1B5E20", "warn":    "#E65100", "error":   "#B71C1C",
}
_LOG_ICONS = {
    "info": "    ", "step": " >> ", "found": " +  ",
    "success": " ✓  ", "warn": " !  ", "error": " ✗  ",
}
_LOG_BOLD = {"step", "success", "error"}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _tmp_path(suffix: str) -> str:
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    return path


def _safe(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _metric_card(val: str, lbl: str, color: str, big: bool = True) -> str:
    num_style = "" if big else (
        "font-size:0.95rem;line-height:1.25;"
        "word-break:break-word;white-space:pre-line;"
    )
    return (
        f'<div class="metric-card" '
        f'style="color:{color};border-top:3px solid {color};">'
        f'<div class="num" style="{num_style}">{_safe(val)}</div>'
        f'<div class="lbl">{lbl}</div>'
        f'</div>'
    )


def _render_log(entries: list) -> str:
    if not entries:
        return (
            '<div class="log-box">'
            '<span style="color:#999;">No activity yet.</span></div>'
        )
    lines = []
    for ts, level, msg in entries:
        color = _LOG_COLORS.get(level, "#999")
        icon  = _LOG_ICONS.get(level, "    ")
        bold  = "font-weight:700;" if level in _LOG_BOLD else ""
        lines.append(
            f'<span style="color:#BBBBBB;font-size:0.72rem;">{ts}</span>'
            f'<span style="color:{color};{bold}">{icon}{_safe(msg)}</span><br>'
        )
    return f'<div class="log-box">{"".join(lines)}</div>'


def _folder_for(unit: dict) -> str:
    """Return 'Assessment' or 'Re-Assessment' based on report type."""
    rt = (unit.get("report_type") or "").strip()
    return "Re-Assessment" if rt.lower().startswith("re") else "Assessment"


def _safe_folder(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", (name or "Unknown").strip()) or "Unknown"


# ── Logo library (persists on disk across sessions) ───────────────────────────
_APP_DIR      = os.path.dirname(os.path.abspath(__file__))
_LOGO_DIR     = os.path.join(_APP_DIR, "saved_logos")
_DEFAULT_LOGO = os.path.join(_APP_DIR, "rvnp_logo.png")
_LOGO_EXTS    = (".png", ".jpg", ".jpeg")


def _logo_label(filename: str) -> str:
    """'greenfield_high.png' → 'Greenfield High'."""
    stem   = os.path.splitext(os.path.basename(filename))[0]
    pretty = re.sub(r"[_-]+", " ", stem).strip().title()
    return pretty or os.path.basename(filename)


def _list_saved_logos() -> list:
    """[(label, path), ...] for every logo in the saved library, A→Z."""
    os.makedirs(_LOGO_DIR, exist_ok=True)
    return [
        (_logo_label(fn), os.path.join(_LOGO_DIR, fn))
        for fn in sorted(os.listdir(_LOGO_DIR))
        if fn.lower().endswith(_LOGO_EXTS)
    ]


def _save_logo(filename: str, blob: bytes) -> str:
    """Persist an uploaded logo to the library; returns its saved path."""
    os.makedirs(_LOGO_DIR, exist_ok=True)
    safe = re.sub(r'[\\/:*?"<>|]', "_", filename).strip() or "logo.png"
    if not safe.lower().endswith(_LOGO_EXTS):
        safe += ".png"
    path = os.path.join(_LOGO_DIR, safe)
    with open(path, "wb") as fh:
        fh.write(blob)
    return path


def _gen_zip(data: dict, logo_path: str = None) -> bytes:
    # Plan each file's folder path + filename up front so we can detect
    # collisions (two units sharing a name within the same course folder).
    planned = [
        (f'{_folder_for(unit)}/'
         f'{_safe_folder(unit.get("course_name") or "Unknown Course")}',
         filename, file_bytes)
        for unit, (filename, file_bytes) in zip(
            data["units"], build_marksheet_per_unit(data, logo_path=logo_path)
        )
    ]

    # Number filenames that repeat within a folder, compared case-INSENSITIVELY:
    # Windows/macOS treat "Name.xlsx" and "name.xlsx" as the same file, so a
    # second one would overwrite the first on extraction. Collisions become
    # "Name (1).xlsx" / "Name (2).xlsx"; unique names stay as-is.
    totals = Counter(f"{path}/{name}".lower() for path, name, _ in planned)
    seen = {}

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path, name, file_bytes in planned:
            full = f"{path}/{name}".lower()
            if totals[full] > 1:
                seen[full] = seen.get(full, 0) + 1
                stem, ext = os.path.splitext(name)
                name = f"{stem} ({seen[full]}){ext}"
            zf.writestr(f"{path}/{name}", file_bytes)
    return buf.getvalue()


def _gen_summative_zip(data: dict) -> tuple[bytes, int]:
    """ZIP of summative moderated-practical marks sheets, one per unit, split
    into Assessment/Re-Assessment folders like the marksheet ZIP.  The sheet
    carries the fixed CDACC logo, so no school logo is threaded.
    Returns (zip_bytes, n_units)."""
    planned = [
        (f'{_folder_for(unit)}/'
         f'{_safe_folder(unit.get("course_name") or "Unknown Course")}',
         filename, file_bytes)
        for unit, (filename, file_bytes) in zip(
            data["units"], build_summative_per_unit(data)
        )
    ]

    totals = Counter(f"{path}/{name}".lower() for path, name, _ in planned)
    seen = {}

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path, name, file_bytes in planned:
            full = f"{path}/{name}".lower()
            if totals[full] > 1:
                seen[full] = seen.get(full, 0) + 1
                stem, ext = os.path.splitext(name)
                name = f"{stem} ({seen[full]}){ext}"
            zf.writestr(f"{path}/{name}", file_bytes)
    return buf.getvalue(), len(planned)


def _gen_class_zip(data: dict, logo_path: str = None) -> tuple[bytes, int]:
    """ZIP of per-class registration forms.  Returns (zip_bytes, n_classes).
    Class names are unique per course by construction; the numbering guard
    below only catches pathological sanitised-filename collisions."""
    files = build_class_forms(data, logo_path=logo_path)
    totals = Counter(name.lower() for name, _ in files)
    seen = {}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, file_bytes in files:
            key = name.lower()
            if totals[key] > 1:
                seen[key] = seen.get(key, 0) + 1
                stem, ext = os.path.splitext(name)
                name = f"{stem} ({seen[key]}){ext}"
            zf.writestr(name, file_bytes)
    return buf.getvalue(), len(files)


def _merge_data(data_list: list) -> dict:
    if not data_list:
        return {}
    if len(data_list) == 1:
        return data_list[0]

    def _uniq(key: str) -> str:
        vals = list(dict.fromkeys(
            d.get(key, "").strip() for d in data_list
            if d.get(key, "").strip()
        ))
        return vals[0] if len(vals) == 1 else " / ".join(vals[:3]) if vals else ""

    all_units = []
    for d in data_list:
        all_units.extend(d.get("units", []))

    return {
        "centre_name":  _uniq("centre_name"),
        "centre_code":  _uniq("centre_code"),
        "course_name":  _uniq("course_name"),
        "course_level": _uniq("course_level"),
        "series":       _uniq("series"),
        "units":        all_units,
        "unit_count":   len(all_units),
    }


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="app-header">
  <h1>📝&nbsp;&nbsp;Nominal Roll &rarr; Marksheet</h1>
  <span class="ver">{_VERSION}</span>
</div>
""", unsafe_allow_html=True)

# ── File uploader ─────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Upload nominal roll PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="Select one or more TVET CDACC Nominal Roll PDFs.",
)

if not uploaded_files:
    st.info(
        "Upload one or more TVET CDACC Nominal Roll PDFs to begin.  "
        "The app will extract candidate lists and generate ready-to-fill "
        "marksheets.",
        icon="📂",
    )
    st.stop()

# ── Incremental extraction ────────────────────────────────────────────────────
current_map: dict = {f.name: f for f in uploaded_files}
processed:   dict = st.session_state.get("processed_files", {})
all_logs:    list = st.session_state.get("log_entries", [])

removed = [name for name in list(processed.keys()) if name not in current_map]
added   = [f for name, f in current_map.items() if name not in processed]

for name in removed:
    del processed[name]

for uploaded in added:
    tmp = _tmp_path(".pdf")
    file_logs: list = []
    try:
        with open(tmp, "wb") as fh:
            fh.write(uploaded.getvalue())

        with st.status(f"🔍  Reading {uploaded.name}…", expanded=True) as status:
            prog = st.progress(0, text="Initialising…")

            def on_log(level: str, msg: str,
                       _logs: list = file_logs) -> None:
                ts = datetime.now().strftime("%H:%M:%S")
                _logs.append((ts, level, msg))
                _icons = {
                    "step": "🔍", "found": "📋", "success": "✅",
                    "warn": "⚠️", "error": "❌", "info": "·",
                }
                st.write(f"{_icons.get(level, '·')}  {msg}")

            def on_progress(cur: int, total: int,
                            _p=prog) -> None:
                _p.progress(cur / total, text=f"Page {cur} of {total}")

            data = extract(tmp, on_log=on_log, on_progress=on_progress)
            prog.progress(1.0, text="Complete ✓")
            status.update(
                label=f"✅  {uploaded.name} — extraction complete!",
                state="complete",
                expanded=False,
            )

        processed[uploaded.name] = data
        all_logs.extend(file_logs)

    except Exception as exc:
        st.error(f"❌  {uploaded.name}: {exc}")
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)

if removed or added:
    st.session_state.processed_files = processed
    st.session_state.log_entries     = all_logs
    st.session_state.extraction_done = bool(processed)
    st.session_state.data            = _merge_data(list(processed.values()))
    for key in list(st.session_state.keys()):
        if key.startswith("_cache_"):
            del st.session_state[key]

if not processed:
    st.warning("No files were processed successfully.")
    st.stop()

# ── Working values ────────────────────────────────────────────────────────────
data        = st.session_state.data
log_entries = st.session_state.get("log_entries", [])
extr_done   = st.session_state.get("extraction_done", False)

total    = sum(u["candidate_count"] for u in data["units"])
assess   = sum(
    u["candidate_count"] for u in data["units"]
    if (u.get("report_type") or "").lower().startswith("assessment")
)
reassess = total - assess

stem_default = (
    os.path.splitext(uploaded_files[0].name)[0]
    if len(uploaded_files) == 1
    else "nominal_roll"
)

# ── Info cards ────────────────────────────────────────────────────────────────
st.markdown('<p class="sec-lbl">Extracted Information</p>',
            unsafe_allow_html=True)
st.markdown(
    '<div class="metrics-row">'
    + _metric_card(data.get("centre_name") or "—", "Centre Name",  _C["primary"], big=False)
    + _metric_card(data.get("centre_code") or "—", "Centre Code",  _C["info"],    big=False)
    + _metric_card(data.get("course_name") or "—", "Course",       _C["success"], big=False)
    + _metric_card(data.get("series")      or "—", "Exam Series",  _C["warning"], big=False)
    + '</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="metrics-row">'
    + _metric_card(str(data["unit_count"]), "Units",                          _C["primary"])
    + _metric_card(str(total),              "Total<br>Registrations",         _C["info"])
    + _metric_card(str(assess),             "Assessment<br>Registrations",    _C["success"])
    + _metric_card(str(reassess),           "Re-Assessment<br>Registrations", _C["danger"])
    + '</div>',
    unsafe_allow_html=True,
)

# ── Activity log ──────────────────────────────────────────────────────────────
with st.container(border=True):
    n_files = len(processed)
    badge = (
        f'<span class="badge-done">Complete ✓  '
        f'({n_files} file{"s" if n_files != 1 else ""})</span>'
        if extr_done else
        '<span class="badge-wait">Waiting…</span>'
    )
    st.markdown(
        f'<div class="act-header">'
        f'<span class="act-title">⚡ Activity</span>{badge}</div>',
        unsafe_allow_html=True,
    )
    st.progress(1.0 if extr_done else 0.0)
    st.markdown(_render_log(log_entries), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── School logo ───────────────────────────────────────────────────────────────
st.markdown('<p class="sec-lbl">School Logo</p>', unsafe_allow_html=True)

with st.container(border=True):
    logo_up_col, logo_pick_col = st.columns([1, 1], gap="large")

    # ── Upload a new logo (saved to the library for next time) ─────────────────
    with logo_up_col:
        st.markdown('<p class="export-title">⬆ Upload a logo</p>',
                    unsafe_allow_html=True)
        st.markdown(
            '<p class="export-caption">PNG or JPG. It is saved to your logo '
            'library, so next time just pick it from the list — no re-upload.</p>',
            unsafe_allow_html=True,
        )
        _new_logo = st.file_uploader(
            "Upload school logo", type=["png", "jpg", "jpeg"],
            key="logo_upload", label_visibility="collapsed",
        )
        if _new_logo is not None:
            _sig = f"{_new_logo.name}:{_new_logo.size}"
            if st.session_state.get("_logo_saved_sig") != _sig:
                _saved = _save_logo(_new_logo.name, _new_logo.getvalue())
                st.session_state["_logo_saved_sig"] = _sig
                st.session_state["logo_choice"]     = _logo_label(_saved)
                st.success(f"Saved “{_logo_label(_saved)}” to your logo library.")

    # ── Choose from the saved library ─────────────────────────────────────────
    with logo_pick_col:
        st.markdown('<p class="export-title">🏫 Choose logo</p>',
                    unsafe_allow_html=True)

        _options  = [("Default (RVNP)", _DEFAULT_LOGO)] + _list_saved_logos()
        _by_label = {lbl: pth for lbl, pth in _options}
        _labels   = list(_by_label.keys())

        if st.session_state.get("logo_choice") not in _labels:
            st.session_state["logo_choice"] = _labels[0]

        _choice   = st.selectbox(
            "Choose logo", _labels, key="logo_choice",
            label_visibility="collapsed",
        )
        logo_path = _by_label.get(_choice, _DEFAULT_LOGO)

        if os.path.exists(logo_path):
            st.image(logo_path, width=120, caption=_choice)
        else:
            st.caption("⚠ Logo file not found — the default will be used.")
            logo_path = _DEFAULT_LOGO

# Rebuild the exports whenever the chosen logo changes.
if st.session_state.get("_active_logo") != logo_path:
    for _k in ("_cache_zip", "_cache_classzip"):
        st.session_state.pop(_k, None)
    st.session_state["_active_logo"] = logo_path

st.markdown("<br>", unsafe_allow_html=True)

# ── Export ────────────────────────────────────────────────────────────────────
st.markdown('<p class="sec-lbl">Export Marksheets</p>', unsafe_allow_html=True)

with st.container(border=True):
    zip_col, summ_col, cls_col = st.columns([1, 1, 1], gap="large")

    # ── ZIP: one xlsx per unit ────────────────────────────────────────────────
    with zip_col:
        st.markdown('<p class="export-title">📦 ZIP — One file per unit</p>',
                    unsafe_allow_html=True)
        st.markdown(
            '<p class="export-caption">Each unit gets its own .xlsx marksheet '
            'with candidate list and empty marks columns.</p>',
            unsafe_allow_html=True,
        )
        _zk = "_cache_zip"
        if _zk not in st.session_state:
            with st.spinner("Building marksheets…"):
                st.session_state[_zk] = _gen_zip(data, logo_path)

        st.caption(
            f"{data['unit_count']} file{'s' if data['unit_count'] != 1 else ''} "
            f"inside — named after each unit."
        )
        st.markdown('<p class="fn-label">Save as</p>', unsafe_allow_html=True)
        zip_fn = st.text_input(
            "ZIP filename", value=f"{stem_default}_marksheets",
            key="fn_zip", label_visibility="collapsed",
        )
        st.download_button(
            label="⬇  Download ZIP",
            data=st.session_state[_zk],
            file_name=f"{zip_fn.strip() or stem_default}.zip",
            mime="application/zip",
            width="stretch",
            type="primary",
        )

    # ── Summative moderated-practical sheets: one xlsx per unit, zipped ───────
    with summ_col:
        st.markdown('<p class="export-title">📑 Summative — ZIP</p>',
                    unsafe_allow_html=True)
        st.markdown(
            '<p class="export-caption">CDACC Summative Assessment Moderated '
            'Practical Marks Sheet per unit — internal, external & moderated '
            'marks columns. Carries the CDACC logo.</p>',
            unsafe_allow_html=True,
        )
        _sk = "_cache_summzip"
        if _sk not in st.session_state:
            with st.spinner("Building summative sheets…"):
                st.session_state[_sk] = _gen_summative_zip(data)
        _szip, _nsumm = st.session_state[_sk]

        st.caption(
            f"{_nsumm} file{'s' if _nsumm != 1 else ''} inside — one per unit."
        )
        st.markdown('<p class="fn-label">Save as</p>', unsafe_allow_html=True)
        summ_fn = st.text_input(
            "Summative filename", value=f"{stem_default}_summative",
            key="fn_summ", label_visibility="collapsed",
        )
        st.download_button(
            label="⬇  Download Summative",
            data=_szip,
            file_name=f"{summ_fn.strip() or stem_default}.zip",
            mime="application/zip",
            width="stretch",
            type="primary",
        )

    # ── Class registration forms: one xlsx per class, zipped ─────────────────
    with cls_col:
        st.markdown('<p class="export-title">🏫 Class Forms — ZIP</p>',
                    unsafe_allow_html=True)
        st.markdown(
            '<p class="export-caption">Classes are read from admission '
            'numbers (the intake code after the first “/”, e.g. 24S). One '
            'registration form per class, named after its course (e.g. '
            '“ICT L6 Class 24S”).</p>',
            unsafe_allow_html=True,
        )
        _ck = "_cache_classzip"
        if _ck not in st.session_state:
            with st.spinner("Building class forms…"):
                st.session_state[_ck] = _gen_class_zip(data, logo_path)
        _czip, _ncls = st.session_state[_ck]

        st.caption(f"{_ncls} class form{'s' if _ncls != 1 else ''} inside.")
        st.markdown('<p class="fn-label">Save as</p>', unsafe_allow_html=True)
        cls_fn = st.text_input(
            "Class forms filename", value=f"{stem_default}_class_forms",
            key="fn_classzip", label_visibility="collapsed",
        )
        st.download_button(
            label="⬇  Download Class Forms",
            data=_czip,
            file_name=f"{cls_fn.strip() or stem_default}.zip",
            mime="application/zip",
            width="stretch",
            type="primary",
        )

# ── Unit breakdown table ──────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown('<p class="sec-lbl">Unit Breakdown</p>', unsafe_allow_html=True)

rows = []
for u in data["units"]:
    rows.append({
        "Unit Name":   u["unit_name"],
        "Report Type": u.get("report_type") or "",
        "Candidates":  u["candidate_count"],
    })

st.dataframe(
    rows,
    width="stretch",
    hide_index=True,
    column_config={
        "Unit Name":   st.column_config.TextColumn(width="large"),
        "Report Type": st.column_config.TextColumn(width="medium"),
        "Candidates":  st.column_config.NumberColumn(width="small"),
    },
)
