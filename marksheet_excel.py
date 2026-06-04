#!/usr/bin/env python3
"""
marksheet_excel.py
==================
Build TVET CDACC marksheet workbooks matching the official sample template.

Adapted for the Nominal Roll Extractor:
  • Column C (Center Admission No.) is populated from cand["admission_no"].
  • Institution name (row 4) is read from data["centre_name"] with a fallback
    to the _CENTRE_NAME constant below.

Row layout:
  1-3   Logo (30 pt each)
  4     Institution name  — bold 13 pt, centred
  5     Department        — bold 13 pt, centred
  6     CAMS/2            — bold 12 pt, left
  7     Blue section title — bold 12 pt, centred
  8     Assessment Center Code | Name  (28 pt, top-left)
  9     Course Code | Course Title     (28 pt, top-left)
  10    Unit Code | Unit Title | Series (28 pt, top-left)
  11    Blank spacing row (8 pt)
  12    Column headers — repeated on every printed page (28 pt)
  13    Sub-headers CAT / PRAC (28 pt)
  14+   Data rows (28 pt)

AVG: only calculated when all three marks are entered (COUNT = 3).
Marks cells (E-G, I-K per data row) are the only editable cells.
Sheet password: 2026.
Print: A4 landscape, fit all columns on one page, centred horizontally.
Footer: "Page X of Y" — Times New Roman 12, centred.
Rows 12:13 repeated at top of every page.
"""

import io
import os
import re

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, InlineFont, TextBlock
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import (
    AnchorMarker, OneCellAnchor, XDRPositiveSize2D,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side

# ── Paths ─────────────────────────────────────────────────────────────────────

_HERE      = os.path.dirname(os.path.abspath(__file__))
_LOGO_PATH = os.path.join(_HERE, "rvnp_logo.png")
_EMU       = 9525   # 1 pixel → EMUs

# ── Constants ─────────────────────────────────────────────────────────────────

_CENTRE_NAME = "RIFT VALLEY NATIONAL POLYTECHNIC"
_DEPT_NAME   = "ICT DEPARTMENT"
_FONT        = "Times New Roman"
_BLUE        = "1F4E79"

_NONE = Side(style=None)
_THIN = Side(style="thin", color="000000")

_BORDER_FULL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CAT_FILL    = PatternFill("solid", fgColor="D9D9D9")
_PRAC_FILL   = PatternFill("solid", fgColor="FBE4D5")
_UNLOCKED    = Protection(locked=False)

_CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LFT = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_LFT_T = Alignment(horizontal="left", vertical="top",    wrap_text=True)
_LFT_B = Alignment(horizontal="left", vertical="bottom", wrap_text=True)


def _f(bold=False, size=12, color="000000"):
    return Font(name=_FONT, bold=bold, size=size, color=color)


def _rich(label: str, value: str, size: int = 12) -> CellRichText:
    return CellRichText(
        TextBlock(InlineFont(b=True,  rFont=_FONT, sz=size), label),
        TextBlock(InlineFont(b=False, rFont=_FONT, sz=size), value),
    )


# ── Helpers ───────────────────────────────────────────────────────────────────

def _lvl(course_level: str) -> str:
    s = (course_level or "").strip()
    return s if s.lower().startswith("level") else (f"Level {s}" if s else "")


def _safe_filename(unit_name: str, unit_code: str) -> str:
    raw   = f"{unit_name}_{unit_code}" if unit_code else unit_name
    clean = re.sub(r'[\\/:*?"<>|]', "_", raw).strip(". ")
    return clean[:200] or "marksheet"


def _safe_sheet_name(name: str, used: set) -> str:
    clean = re.sub(r"[:\\/?*\[\]]", "-", name)[:31].strip() or "Sheet"
    cand, n = clean, 1
    while cand.lower() in used:
        suffix = f"_{n}"
        cand = clean[: 31 - len(suffix)] + suffix
        n += 1
    used.add(cand.lower())
    return cand


def _set(ws, coord, value=None, font=None, align=None, fill=None, border=None):
    c = ws[coord]
    if value  is not None: c.value     = value
    if font   is not None: c.font      = font
    if align  is not None: c.alignment = align
    if fill   is not None: c.fill      = fill
    if border is not None: c.border    = border


def _border_range(ws, r1, c1, r2, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = _BORDER_FULL


# ── Sheet builder ─────────────────────────────────────────────────────────────

def _build_unit_sheet(ws, data: dict, unit: dict):
    centre_code  = data.get("centre_code") or ""
    centre_name  = data.get("centre_name") or _CENTRE_NAME
    course_name  = (unit.get("course_name") or data.get("course_name") or "")
    course_level = (unit.get("course_level") or data.get("course_level") or "")
    unit_code    = unit.get("unit_code") or ""
    unit_name    = unit["unit_name"]
    series       = data.get("series") or ""
    candidates   = sorted(unit["candidates"], key=lambda c: c["sn"])

    lvl         = _lvl(course_level)
    course_full = f"{course_name}  {lvl}".strip() if lvl else course_name

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 4.55
    ws.column_dimensions["B"].width = 25.22
    ws.column_dimensions["C"].width = 22.0
    ws.column_dimensions["D"].width = 28.66
    for col in "EFGHIJKL":
        ws.column_dimensions[col].width = 8.0

    # ── Row heights ───────────────────────────────────────────────────────────
    for r in (1, 2, 3):
        ws.row_dimensions[r].height = 30.0
    for r in (8, 9, 10):
        ws.row_dimensions[r].height = 28.0
    ws.row_dimensions[11].height = 14.0
    for r in (12, 13):
        ws.row_dimensions[r].height = 28.0

    # ── Rows 1-3 : logo centred across A-L ────────────────────────────────────
    if os.path.exists(_LOGO_PATH):
        xl_img  = XLImage(_LOGO_PATH)
        logo_px = 100
        marker  = AnchorMarker(col=3, colOff=87 * _EMU, row=0, rowOff=10 * _EMU)
        size    = XDRPositiveSize2D(cx=logo_px * _EMU, cy=logo_px * _EMU)
        xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
        ws.add_image(xl_img)

    # ── Row 4 : institution name ───────────────────────────────────────────────
    ws.merge_cells("A4:L4")
    _set(ws, "A4", value=centre_name, font=_f(bold=True, size=13), align=_CTR)

    # ── Row 5 : department ────────────────────────────────────────────────────
    ws.merge_cells("A5:L5")
    _set(ws, "A5", value=_DEPT_NAME, font=_f(bold=True, size=13), align=_CTR)

    # ── Row 6 : reference ─────────────────────────────────────────────────────
    ws.merge_cells("A6:L6")
    _set(ws, "A6", value="CAMS/2", font=_f(bold=True, size=12), align=_LFT)

    # ── Row 7 : blue section title ────────────────────────────────────────────
    ws.merge_cells("A7:L7")
    _set(ws, "A7",
         value="CONTINUOUS ASSESSMENT MARKS SHEETS PER UNIT OF COMPETENCY",
         font=_f(bold=True, size=12, color=_BLUE), align=_CTR)

    # ── Row 8 : Assessment Center Code | Name ─────────────────────────────────
    ws.merge_cells("A8:D8")
    ws.merge_cells("E8:L8")
    ws["A8"].value     = _rich("Assessment Center Code:   ", centre_code)
    ws["A8"].alignment = _LFT
    ws["E8"].value     = _rich("Assessment Center Name:   ", centre_name)
    ws["E8"].alignment = _LFT

    # ── Row 9 : Course Title | Course Code ────────────────────────────────────
    ws.merge_cells("A9:D9")
    ws.merge_cells("E9:L9")
    ws["A9"].value     = _rich("Course Title:   ", course_full)
    ws["A9"].alignment = _LFT
    ws["E9"].value     = _rich("Course Code:   ", "")
    ws["E9"].alignment = _LFT

    # ── Row 10 : Unit Title | Unit Code | Assessment Series ───────────────────
    ws.merge_cells("A10:D10")
    ws.merge_cells("E10:H10")
    ws.merge_cells("I10:L10")
    ws["A10"].value     = _rich("Unit Title:   ", unit_name)
    ws["A10"].alignment = _LFT
    ws["E10"].value     = _rich("Unit Code:   ", unit_code)
    ws["E10"].alignment = _LFT
    ws["I10"].value     = _rich("Assessment Series:   ", series)
    ws["I10"].alignment = _LFT

    # ── Row 11 : blank spacing row ────────────────────────────────────────────
    for c in range(1, 13):
        ws.cell(row=11, column=c).border = Border(
            left=_NONE, right=_NONE, top=_NONE, bottom=_THIN
        )
    for c in (5, 7):
        ws.cell(row=11, column=c).alignment = _CTR

    # ── Row 12 : main column headers ──────────────────────────────────────────
    for merge in ("A12:A13", "B12:B13", "C12:C13", "D12:D13",
                  "E12:H12", "I12:L12"):
        ws.merge_cells(merge)

    for coord, label in (
        ("A12", "S/N"),
        ("B12", "Candidate's Reg Code"),
        ("C12", "Center Admission No."),
        ("D12", "Candidate's Name"),
        ("E12", "Oral/Theory Marks (100%)"),
        ("I12", "Practical Marks (100%)"),
    ):
        _set(ws, coord, value=label, font=_f(bold=True, size=12), align=_CTR)

    # ── Row 13 : sub-headers ──────────────────────────────────────────────────
    for coord, label, fill in (
        ("E13", "CAT 1",  _CAT_FILL),
        ("F13", "CAT 2",  _CAT_FILL),
        ("G13", "CAT 3",  _CAT_FILL),
        ("H13", "AVG",    _CAT_FILL),
        ("I13", "PRAC 1", _PRAC_FILL),
        ("J13", "PRAC 2", _PRAC_FILL),
        ("K13", "PRAC 3", _PRAC_FILL),
        ("L13", "AVG",    _PRAC_FILL),
    ):
        _set(ws, coord, value=label,
             font=_f(bold=True, size=12), align=_CTR, fill=fill)

    _border_range(ws, 12, 1, 13, 12)

    # ── Data rows from row 14 ─────────────────────────────────────────────────
    for i, cand in enumerate(candidates):
        r = 14 + i
        ws.row_dimensions[r].height = 28.0

        ws.cell(row=r, column=1, value=cand["sn"]).alignment           = _CTR
        ws.cell(row=r, column=2, value=cand["reg_no"]).alignment       = _LFT
        ws.cell(row=r, column=3,
                value=cand.get("admission_no", "")).alignment          = _LFT
        ws.cell(row=r, column=4, value=cand["name"]).alignment         = _LFT

        cat_avg           = ws.cell(row=r, column=8)
        cat_avg.value     = f'=IF(COUNT(E{r}:G{r})=3,AVERAGE(E{r}:G{r}),"")'
        cat_avg.alignment = _CTR

        prac_avg           = ws.cell(row=r, column=12)
        prac_avg.value     = f'=IF(COUNT(I{r}:K{r})=3,AVERAGE(I{r}:K{r}),"")'
        prac_avg.alignment = _CTR

        for c in range(1, 13):
            ws.cell(row=r, column=c).font   = _f(size=12)
            ws.cell(row=r, column=c).border = _BORDER_FULL

        for c in (5, 6, 7, 9, 10, 11):
            ws.cell(row=r, column=c).alignment  = _CTR
            ws.cell(row=r, column=c).protection = _UNLOCKED

    # ── Footer rows ───────────────────────────────────────────────────────────
    last         = 13 + len(candidates)
    prepared_row = last + 2
    approved_row = last + 3

    ws.row_dimensions[last + 1].height   = 28.0
    ws.row_dimensions[prepared_row].height = 28.0
    ws.row_dimensions[approved_row].height = 28.0

    for merge, col, label in (
        (f"A{prepared_row}:C{prepared_row}", 1, "Prepared by:"),
        (f"D{prepared_row}:G{prepared_row}", 4, "Signature:"),
        (f"H{prepared_row}:L{prepared_row}", 8, "Date:"),
    ):
        ws.merge_cells(merge)
        cell           = ws.cell(row=prepared_row, column=col)
        cell.value     = label
        cell.font      = _f(bold=True, size=12)
        cell.alignment = _LFT

    for merge, col, label in (
        (f"A{approved_row}:C{approved_row}", 1, "Approved by:"),
        (f"D{approved_row}:G{approved_row}", 4, "Signature:"),
        (f"H{approved_row}:L{approved_row}", 8, "Date:"),
    ):
        ws.merge_cells(merge)
        cell           = ws.cell(row=approved_row, column=col)
        cell.value     = label
        cell.font      = _f(bold=True, size=12)
        cell.alignment = _LFT_B

    # ── Print settings ────────────────────────────────────────────────────────
    ws.print_area = f"A1:L{approved_row}"
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
    ws.print_title_rows = "12:13"
    ws.protection.sheet               = True
    ws.protection.password            = "2026"
    ws.protection.selectLockedCells   = False
    ws.protection.selectUnlockedCells = False


# ── Public entry points ───────────────────────────────────────────────────────

def build_marksheet_per_unit(data: dict) -> list[tuple[str, bytes]]:
    """One .xlsx per unit.  Returns [(filename, file_bytes), ...]."""
    results: list[tuple[str, bytes]] = []
    for unit in data["units"]:
        wb = Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet(
            _safe_sheet_name(
                re.sub(r"[/\\:*?\[\]]", "-", unit["unit_name"])[:31], set()
            )
        )
        _build_unit_sheet(ws, data, unit)
        buf = io.BytesIO()
        wb.save(buf)
        fname = _safe_filename(unit["unit_name"], unit.get("unit_code", "")) + ".xlsx"
        results.append((fname, buf.getvalue()))
    return results


def build_marksheet_workbook(data: dict) -> bytes:
    """All units as separate sheets in one workbook.  Returns file bytes."""
    wb, used = Workbook(), set()
    wb.remove(wb.active)
    for unit in data["units"]:
        ws = wb.create_sheet(
            _safe_sheet_name(
                re.sub(r"[/\\:*?\[\]]", "-", unit["unit_name"])[:31], used
            )
        )
        _build_unit_sheet(ws, data, unit)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
