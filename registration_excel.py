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
out as an Excel table so the data can be filtered/sorted.  Re-Assessment
units are excluded entirely — the form registers first-attempt candidates
only.

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
    _border_range, _f, _rich, _safe_sheet_name, _set,
)

_EXAMINING_BODY = "TVET CDACC"


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
    first-seen order preserved."""
    seen, out = set(), []
    for unit in course["units"]:
        for cand in sorted(unit["candidates"], key=lambda c: c["sn"]):
            key = cand.get("reg_no") or (cand.get("name"), cand.get("admission_no"))
            if key in seen:
                continue
            seen.add(key)
            out.append({
                "name":         cand.get("name", ""),
                "admission_no": cand.get("admission_no", ""),
                "reg_no":       cand.get("reg_no", ""),
                "level":        course["course_level"],
            })
    return out


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


# ── Sheet builder ─────────────────────────────────────────────────────────────

def _build_form_sheet(ws, data: dict, course: dict, table_id: int,
                      logo_path: str = None):
    centre_name = data.get("centre_name") or _CENTRE_NAME
    candidates  = _union_candidates(course)
    units       = _unit_labels(course)

    # ── Column widths ─────────────────────────────────────────────────────────
    for col, width in (("A", 7.0), ("B", 34.0), ("C", 16.0), ("D", 32.0),
                       ("E", 9.0), ("F", 13.0), ("G", 16.0), ("H", 18.0)):
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
    ws.merge_cells("A4:H4")
    _set(ws, "A4", value=centre_name, font=_f(bold=True, size=14), align=_CTR)

    ws.merge_cells("A5:H5")
    _set(ws, "A5", value="ISO 2009:2015 QUALITY MANAGEMENT SYSTEM",
         font=_f(bold=True, size=12), align=_CTR)

    ws.merge_cells("A6:H6")
    _set(ws, "A6", value="SUMMATIVE ASSESSMENT REGISTRATION",
         font=_f(bold=True, size=12, color=_BLUE), align=_CTR)

    ws.row_dimensions[7].height = 8.0

    # ── Rows 8-11 : form fields ───────────────────────────────────────────────
    for row, label, value in (
        (8,  "DEPARTMENT: ",     "_" * 45),
        (9,  "EXAMINING BODY: ", _EXAMINING_BODY),
        (10, "COURSE NAME: ",    course["course_name"]),
        (11, "CLASS NAME: ",     "_" * 45),
    ):
        ws.merge_cells(f"A{row}:H{row}")
        ws.row_dimensions[row].height = 20.0
        _set(ws, f"A{row}", value=_rich(label, value), align=_LFT)

    ws.row_dimensions[12].height = 8.0

    # ── Row 13 : table header ─────────────────────────────────────────────────
    ws.row_dimensions[13].height = 24.0
    for col, label in zip("ABCDEFGH",
                          ("S/N", "NAME", "ADM NO", "REG. NO", "LEVEL",
                           "ASS. FEES", "FEES ARREARS", "REMARKS")):
        _set(ws, f"{col}13", value=label, font=_f(bold=True, size=12), align=_CTR)
    _border_range(ws, 13, 1, 13, 8)

    # ── Rows 14.. : candidates ────────────────────────────────────────────────
    for i, cand in enumerate(candidates, 1):
        r = 13 + i
        ws.row_dimensions[r].height = 18.0
        _set(ws, f"A{r}", value=i,                    font=_f(size=12), align=_CTR)
        _set(ws, f"B{r}", value=cand["name"],         font=_f(size=12), align=_LFT)
        _set(ws, f"C{r}", value=cand["admission_no"], font=_f(size=12), align=_LFT)
        _set(ws, f"D{r}", value=cand["reg_no"],       font=_f(size=12), align=_LFT)
        _set(ws, f"E{r}", value=cand["level"],        font=_f(size=12), align=_CTR)
        _border_range(ws, r, 1, r, 8)

    # Excel table over header + data so the list can be filtered/sorted.
    # Column names are synced from the row-13 cells at save time.
    if candidates:
        ws.add_table(Table(displayName=f"RegCandidates{table_id}",
                           ref=f"A13:H{13 + len(candidates)}"))

    # ── Units registered ──────────────────────────────────────────────────────
    last      = 13 + len(candidates)
    units_row = last + 2
    ws.merge_cells(f"A{units_row}:C{units_row}")
    _set(ws, f"A{units_row}", value="UNITS REGISTERED",
         font=_f(bold=True, size=12), align=_LFT)

    for i, label in enumerate(units, 1):
        r = units_row + i
        ws.row_dimensions[r].height = 18.0
        ws.merge_cells(f"B{r}:H{r}")
        _set(ws, f"A{r}", value=i,     font=_f(size=12), align=_CTR)
        _set(ws, f"B{r}", value=label, font=_f(size=12), align=_LFT)

    # ── Signature footer ──────────────────────────────────────────────────────
    prep_row = units_row + len(units) + 2
    appr_row = prep_row + 3

    for row, who in ((prep_row, "PREPARED BY: "), (appr_row, "APPROVED BY: ")):
        ws.row_dimensions[row].height = 22.0
        for merge, col, label, value in (
            (f"A{row}:C{row}", "A", who,      "_" * 22),
            (f"D{row}:E{row}", "D", "SIGN: ", "_" * 15),
            (f"F{row}:H{row}", "F", "DATE: ", "_" * 15),
        ):
            ws.merge_cells(merge)
            _set(ws, f"{col}{row}", value=_rich(label, value), align=_LFT)

    for row, title in ((prep_row + 1, "DEPARTMENTAL EO"), (appr_row + 1, "HOD")):
        ws.merge_cells(f"A{row}:C{row}")
        _set(ws, f"A{row}", value=title, font=_f(size=11), align=_LFT)

    # ── Print settings ────────────────────────────────────────────────────────
    ws.print_area = f"A1:H{appr_row + 1}"
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
