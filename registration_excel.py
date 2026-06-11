#!/usr/bin/env python3
"""
registration_excel.py
=====================
Build a SUMMATIVE ASSESSMENT REGISTRATION form from extracted nominal-roll
data.  Layout follows the departmental template (ASSESSMENT REGISTRATION
FORM.xlsx) rebuilt in the marksheet house style, hence the imports of
marksheet_excel's private style helpers — both exports stay visually
consistent.

One form per course: candidates are the union across all units (deduplicated
by reg no, first-seen order, renumbered 1..N) and every unit is listed under
UNITS REGISTERED, Re-Assessment units suffixed "(Re-Assessment)".

DEPARTMENT, CLASS NAME, ASS. FEES, FEES ARREARS and REMARKS are left blank
for manual entry; the sheet is intentionally unprotected.
"""

import io
import os
import re

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import (
    AnchorMarker, OneCellAnchor, XDRPositiveSize2D,
)

from marksheet_excel import (
    _BLUE, _BORDER_FULL, _CENTRE_NAME, _CTR, _EMU, _LFT, _LOGO_PATH,
    _border_range, _f, _rich, _set,
)

_EXAMINING_BODY = "TVET CDACC"


# ── Data helpers ──────────────────────────────────────────────────────────────

def _level_num(course_level: str) -> str:
    """'Level 5' → '5'; bare numbers pass through; no digit → raw trimmed."""
    m = re.search(r"\d+", course_level or "")
    return m.group(0) if m else (course_level or "").strip()


def _union_candidates(data: dict) -> list[dict]:
    """Union of candidates across all units, deduplicated by reg no,
    first-seen order preserved."""
    seen, out = set(), []
    for unit in data.get("units", []):
        lvl = _level_num(unit.get("course_level") or data.get("course_level") or "")
        for cand in sorted(unit["candidates"], key=lambda c: c["sn"]):
            key = cand.get("reg_no") or (cand.get("name"), cand.get("admission_no"))
            if key in seen:
                continue
            seen.add(key)
            out.append({
                "name":         cand.get("name", ""),
                "admission_no": cand.get("admission_no", ""),
                "reg_no":       cand.get("reg_no", ""),
                "level":        lvl,
            })
    return out


def _unit_labels(data: dict) -> list[str]:
    """All units in roll order, deduplicated by (report_type, unit_name)."""
    seen, labels = set(), []
    for u in data.get("units", []):
        rt  = (u.get("report_type") or "").strip()
        key = (rt, u["unit_name"])
        if key in seen:
            continue
        seen.add(key)
        suffix = "  (Re-Assessment)" if rt.lower().startswith("re") else ""
        labels.append(f"{u['unit_name']}{suffix}")
    return labels


# ── Sheet builder ─────────────────────────────────────────────────────────────

def _build_form_sheet(ws, data: dict, logo_path: str = None):
    centre_name = data.get("centre_name") or _CENTRE_NAME
    course_name = data.get("course_name") or ""
    candidates  = _union_candidates(data)
    units       = _unit_labels(data)

    # ── Column widths ─────────────────────────────────────────────────────────
    for col, width in (("A", 6.0), ("B", 30.0), ("C", 14.0), ("D", 30.0),
                       ("E", 8.0), ("F", 11.0), ("G", 14.0), ("H", 16.0)):
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
        (10, "COURSE NAME: ",    course_name),
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
    """One consolidated registration form for the whole course.  Returns
    file bytes."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Registration"
    _build_form_sheet(ws, data, logo_path)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
