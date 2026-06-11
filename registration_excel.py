#!/usr/bin/env python3
"""
registration_excel.py
=====================
Build a SUMMATIVE ASSESSMENT REGISTRATION workbook from extracted
nominal-roll data.  Layout follows the departmental template (ASSESSMENT
REGISTRATION FORM.xlsx) rebuilt in the marksheet house style, hence the
imports of marksheet_excel's private style helpers — both exports stay
visually consistent.

One sheet tab per course: candidates are the union across that course's
units (deduplicated by reg no, first-seen order, renumbered 1..N) and laid
out as an Excel table so the data can be filtered/sorted.  Each candidate
row carries the count and names of the units they registered (one name per
line; the column is sized to the longest unit name).  Re-Assessment units
are excluded entirely — the form registers first-attempt candidates only.

DEPARTMENT, CLASS NAME, ASS. FEES, FEES ARREARS and REMARKS are left blank
for manual entry; the sheets are intentionally unprotected.
"""

import io
import os
import re

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import (
    AnchorMarker, OneCellAnchor, XDRPositiveSize2D,
)
from openpyxl.worksheet.table import Table

from marksheet_excel import (
    _BLUE, _BORDER_FULL, _CENTRE_NAME, _CTR, _EMU, _LFT, _LOGO_PATH,
    _border_range, _f, _rich, _safe_filename, _safe_sheet_name, _set,
)

_EXAMINING_BODY = "TVET CDACC"

_HEADERS = ("S/N", "NAME", "ADM NO", "REG. NO", "LEVEL", "UNIT(S)",
            "UNIT(S) REGISTERED NAME(S)", "ASS. FEES", "FEES ARREARS",
            "REMARKS")
_LAST_COL = "J"          # 10 columns, A..J
_UNITS_COL = "G"


# ── Data helpers ──────────────────────────────────────────────────────────────

def _is_reassessment(unit: dict) -> bool:
    return (unit.get("report_type") or "").strip().lower().startswith("re")


def _level_num(course_level: str) -> str:
    """'Level 5' → '5'; bare numbers pass through; no digit → raw trimmed."""
    m = re.search(r"\d+", course_level or "")
    return m.group(0) if m else (course_level or "").strip()


def _courses(data: dict) -> list[dict]:
    """Group assessment units by course (name + level), roll order preserved.
    Re-Assessment units are dropped."""
    groups, order = {}, []
    for u in data.get("units", []):
        if _is_reassessment(u):
            continue
        cname = (u.get("course_name") or data.get("course_name") or "").strip()
        clvl  = _level_num(u.get("course_level") or data.get("course_level") or "")
        key   = (cname.upper(), clvl)
        if key not in groups:
            groups[key] = {"course_name": cname, "course_level": clvl, "units": []}
            order.append(key)
        groups[key]["units"].append(u)
    if not order:   # roll held nothing but re-assessments
        return [{"course_name": (data.get("course_name") or "").strip(),
                 "course_level": _level_num(data.get("course_level") or ""),
                 "units": []}]
    return [groups[k] for k in order]


def _union_candidates(course: dict) -> list[dict]:
    """Union of candidates across the course's units, deduplicated by reg no,
    first-seen order preserved.  Each record accumulates the names of the
    units the candidate appears in."""
    by_key, order = {}, []
    for unit in course["units"]:
        for cand in sorted(unit["candidates"], key=lambda c: c["sn"]):
            key = cand.get("reg_no") or (cand.get("name"), cand.get("admission_no"))
            rec = by_key.get(key)
            if rec is None:
                rec = {
                    "name":         cand.get("name", ""),
                    "admission_no": cand.get("admission_no", ""),
                    "reg_no":       cand.get("reg_no", ""),
                    "level":        course["course_level"],
                    "units":        [],
                }
                by_key[key] = rec
                order.append(key)
            if unit["unit_name"] not in rec["units"]:
                rec["units"].append(unit["unit_name"])
    return [by_key[k] for k in order]


def _unit_labels(course: dict) -> list[str]:
    """The course's unit names in roll order, deduplicated."""
    seen, labels = set(), []
    for u in course["units"]:
        if u["unit_name"] in seen:
            continue
        seen.add(u["unit_name"])
        labels.append(u["unit_name"])
    return labels


def _sheet_title(course: dict) -> str:
    """Course name (+ level) trimmed so the level survives the 31-char cap."""
    cname, clvl = course["course_name"] or "Registration", course["course_level"]
    if not clvl:
        return cname
    suffix = f" L{clvl}"
    return cname[: 31 - len(suffix)].rstrip() + suffix


_CLASS_CODE_RE = re.compile(r"/\s*(\d{2}[A-Za-z])")


def _class_code(admission_no: str) -> str:
    """Class identifier from an admission number: the intake token after the
    first '/', e.g. '1234/24S' → '24S'.  Empty when the number has none."""
    m = _CLASS_CODE_RE.search(admission_no or "")
    return m.group(1).upper() if m else ""


def _class_groups(course: dict) -> list[dict]:
    """Split a course's candidates into classes identified by the intake code
    in their admission numbers (24S, 25M, …).  Candidates whose admission
    number carries no code fall into UNGROUPED.  Each class gets a name
    unique to its course; classes are ordered by code (UNGROUPED last)."""
    groups = {}
    for cand in _union_candidates(course):
        code = _class_code(cand["admission_no"]) or "UNGROUPED"
        groups.setdefault(code, []).append(cand)

    cname, clvl = course["course_name"], course["course_level"]
    stem = f"{cname} L{clvl}" if clvl else cname
    out = []
    for code in sorted(groups, key=lambda c: (c == "UNGROUPED", c)):
        members    = groups[code]
        unit_names = {u for m in members for u in m["units"]}
        out.append({
            "class_name": f"{stem} CLASS {code}".strip(),
            "short_name": f"CLASS {code}",
            "course": {
                "course_name":  cname,
                "course_level": clvl,
                "units": [u for u in course["units"] if u["unit_name"] in unit_names],
            },
            "candidates": members,
        })
    return out


# ── Sheet builder ─────────────────────────────────────────────────────────────

def _build_form_sheet(ws, data: dict, course: dict, table_id: int,
                      logo_path: str = None, candidates: list = None,
                      class_name: str = ""):
    centre_name = data.get("centre_name") or _CENTRE_NAME
    if candidates is None:
        candidates = _union_candidates(course)
    units       = _unit_labels(course)

    # ── Column widths ─────────────────────────────────────────────────────────
    longest_unit = max((len(u) for u in units), default=0)
    for col, width in (("A", 7.0), ("B", 34.0), ("C", 20.0), ("D", 32.0),
                       ("E", 9.0), ("F", 9.0),
                       (_UNITS_COL, max(18.0, longest_unit + 4.0)),
                       ("H", 13.0), ("I", 16.0), ("J", 18.0)):
        ws.column_dimensions[col].width = width

    # ── Rows 1-3 : logo ───────────────────────────────────────────────────────
    for r in (1, 2, 3):
        ws.row_dimensions[r].height = 30.0
    logo = logo_path or _LOGO_PATH
    if logo and os.path.exists(logo):
        xl_img  = XLImage(logo)
        logo_px = 100
        marker  = AnchorMarker(col=3, colOff=50 * _EMU, row=0, rowOff=10 * _EMU)
        size    = XDRPositiveSize2D(cx=logo_px * _EMU, cy=logo_px * _EMU)
        xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
        ws.add_image(xl_img)

    # ── Rows 4-6 : headings ───────────────────────────────────────────────────
    ws.row_dimensions[4].height = 22.0
    ws.merge_cells(f"A4:{_LAST_COL}4")
    _set(ws, "A4", value=centre_name, font=_f(bold=True, size=14), align=_CTR)

    ws.merge_cells(f"A5:{_LAST_COL}5")
    _set(ws, "A5", value="ISO 2009:2015 QUALITY MANAGEMENT SYSTEM",
         font=_f(bold=True, size=12), align=_CTR)

    ws.merge_cells(f"A6:{_LAST_COL}6")
    _set(ws, "A6", value="SUMMATIVE ASSESSMENT REGISTRATION",
         font=_f(bold=True, size=12, color=_BLUE), align=_CTR)

    ws.row_dimensions[7].height = 8.0

    # ── Rows 8-11 : form fields (blank labels stay bare — no underscores) ────
    for row, label, value in (
        (8,  "DEPARTMENT:",      ""),
        (9,  "EXAMINING BODY: ", _EXAMINING_BODY),
        (10, "COURSE NAME: ",    course["course_name"]),
        (11, "CLASS NAME:" if not class_name else "CLASS NAME: ", class_name),
    ):
        ws.merge_cells(f"A{row}:{_LAST_COL}{row}")
        ws.row_dimensions[row].height = 20.0
        if value:
            _set(ws, f"A{row}", value=_rich(label, value), align=_LFT)
        else:
            _set(ws, f"A{row}", value=label, font=_f(bold=True, size=12),
                 align=_LFT)

    ws.row_dimensions[12].height = 8.0

    # ── Row 13 : table header ─────────────────────────────────────────────────
    ws.row_dimensions[13].height = 24.0
    for col, label in zip("ABCDEFGHIJ", _HEADERS):
        _set(ws, f"{col}13", value=label, font=_f(bold=True, size=12), align=_CTR)
    _border_range(ws, 13, 1, 13, 10)

    # ── Rows 14.. : candidates ────────────────────────────────────────────────
    for i, cand in enumerate(candidates, 1):
        r = 13 + i
        n_units = len(cand["units"])
        ws.row_dimensions[r].height = max(18.0, 16.0 * n_units + 2.0)
        _set(ws, f"A{r}", value=i,                       font=_f(size=12), align=_CTR)
        _set(ws, f"B{r}", value=cand["name"],            font=_f(size=12), align=_LFT)
        _set(ws, f"C{r}", value=cand["admission_no"],    font=_f(size=12), align=_LFT)
        _set(ws, f"D{r}", value=cand["reg_no"],          font=_f(size=12), align=_LFT)
        _set(ws, f"E{r}", value=cand["level"],           font=_f(size=12), align=_CTR)
        _set(ws, f"F{r}", value=n_units,                 font=_f(size=12), align=_CTR)
        _set(ws, f"{_UNITS_COL}{r}", value="\n".join(cand["units"]),
             font=_f(size=12), align=_LFT)
        _border_range(ws, r, 1, r, 10)

    # Excel table over header + data so the list can be filtered/sorted.
    # Column names are synced from the row-13 cells at save time.
    if candidates:
        ws.add_table(Table(displayName=f"RegCandidates{table_id}",
                           ref=f"A13:{_LAST_COL}{13 + len(candidates)}"))

    # ── Units registered ──────────────────────────────────────────────────────
    last      = 13 + len(candidates)
    units_row = last + 2
    ws.merge_cells(f"A{units_row}:C{units_row}")
    _set(ws, f"A{units_row}", value="UNITS REGISTERED",
         font=_f(bold=True, size=12), align=_LFT)

    for i, label in enumerate(units, 1):
        r = units_row + i
        ws.row_dimensions[r].height = 18.0
        ws.merge_cells(f"B{r}:{_LAST_COL}{r}")
        _set(ws, f"A{r}", value=i,     font=_f(size=12), align=_CTR)
        _set(ws, f"B{r}", value=label, font=_f(size=12), align=_LFT)

    # ── Signature footer ──────────────────────────────────────────────────────
    prep_row = units_row + len(units) + 2
    appr_row = prep_row + 3

    for row, who in ((prep_row, "PREPARED BY:"), (appr_row, "APPROVED BY:")):
        ws.row_dimensions[row].height = 22.0
        for merge, col, label in (
            (f"A{row}:C{row}", "A", who),
            (f"D{row}:F{row}", "D", "SIGN:"),
            (f"G{row}:{_LAST_COL}{row}", "G", "DATE:"),
        ):
            ws.merge_cells(merge)
            _set(ws, f"{col}{row}", value=label, font=_f(bold=True, size=12),
                 align=_LFT)

    for row, title in ((prep_row + 1, "DEPARTMENTAL EO"), (appr_row + 1, "HOD")):
        ws.merge_cells(f"A{row}:C{row}")
        _set(ws, f"A{row}", value=title, font=_f(size=11), align=_LFT)

    # ── Print settings ────────────────────────────────────────────────────────
    ws.print_area = f"A1:{_LAST_COL}{appr_row + 1}"
    ws.page_setup.paperSize   = 9
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left   = 0.25
    ws.page_margins.right  = 0.25
    ws.page_margins.top    = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.3
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.center.font = "Times New Roman"
    ws.oddFooter.center.size = 12
    ws.print_title_rows = "13:13"


# ── Public entry point ────────────────────────────────────────────────────────

def build_registration_form(data: dict, logo_path: str = None) -> bytes:
    """One registration sheet per course (Re-Assessment units excluded).
    Returns file bytes."""
    wb, used = Workbook(), set()
    wb.remove(wb.active)
    for i, course in enumerate(_courses(data), 1):
        ws = wb.create_sheet(_safe_sheet_name(_sheet_title(course), used))
        _build_form_sheet(ws, data, course, table_id=i, logo_path=logo_path)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_class_forms(data: dict, logo_path: str = None) -> list[tuple[str, bytes]]:
    """One registration form per class — classes are identified by the intake
    code in each candidate's admission number (e.g. '1234/24S' → class 24S)
    and named after their course (e.g. "ICT L6 CLASS 24S").
    Returns [(filename, file_bytes), ...]."""
    results = []
    for course in _courses(data):
        for grp in _class_groups(course):
            wb = Workbook()
            ws = wb.active
            ws.title = _safe_sheet_name(grp["short_name"], set())
            _build_form_sheet(ws, data, grp["course"], table_id=1,
                              logo_path=logo_path,
                              candidates=grp["candidates"],
                              class_name=grp["class_name"])
            buf = io.BytesIO()
            wb.save(buf)
            results.append((_safe_filename(grp["class_name"], "") + ".xlsx",
                            buf.getvalue()))
    return results
