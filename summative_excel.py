#!/usr/bin/env python3
"""
summative_excel.py
==================
Build TVET CDACC **Summative Assessment Moderated Practical Marks Sheets**
(one per unit of competency) from extracted nominal-roll data.

Layout matches the official CDACC moderated-practical form: candidate list
with Internal Assessor / External Verifier / Moderated marks columns, a Mean
Deviation row, and a three-signatory footer (Internal Assessor, External
Verifier, Assessment Centre Manager).  The sheet always carries the fixed
CDACC logo — this is a national CDACC document, not a school one.

Marks columns (Internal / External / Moderated) and the Mean Deviation cell
are the only editable cells; the rest is protected with password 2026.

Reuses the marksheet house-style helpers so both exports stay consistent.
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
from openpyxl.styles import Border, PatternFill

from marksheet_excel import (
    _BLUE, _BORDER_FULL, _CTR, _EMU, _FONT, _LFT, _NONE, _THIN, _UNLOCKED,
    _border_range, _f, _safe_filename, _safe_sheet_name, _set,
)

# ── Paths / fills ─────────────────────────────────────────────────────────────

_HERE        = os.path.dirname(os.path.abspath(__file__))
_CDACC_LOGO  = os.path.join(_HERE, "cdacc_logo.png")
_COL_FILL    = PatternFill("solid", fgColor="D9D9D9")

# marksheet's _border_range applies a full thin border to every cell in a range
_full_border = _border_range


def _rich(*parts, size: int = 11) -> CellRichText:
    """Build CellRichText from alternating (bold, text) pairs."""
    return CellRichText(*(
        TextBlock(InlineFont(b=bold, rFont=_FONT, sz=size), text)
        for bold, text in parts
    ))


def _outer_border(ws, r1, c1, r2, c2):
    """Thin border on the outer edge of a cell range only."""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = Border(
                left   = _THIN if c == c1 else _NONE,
                right  = _THIN if c == c2 else _NONE,
                top    = _THIN if r == r1 else _NONE,
                bottom = _THIN if r == r2 else _NONE,
            )


def _dots(n: int = 50) -> str:
    return "." * n


# ── Sheet builder ─────────────────────────────────────────────────────────────

def build_summative_sheet(ws, data: dict, unit: dict, logo_path: str = None):
    centre_code  = data.get("centre_code") or ""
    centre_name  = data.get("centre_name") or ""
    course_name  = unit.get("course_name") or data.get("course_name") or ""
    course_level = unit.get("course_level") or data.get("course_level") or ""
    unit_code    = unit.get("unit_code") or ""
    unit_name    = unit["unit_name"]
    series       = data.get("series") or ""
    candidates   = sorted(unit["candidates"], key=lambda c: c["sn"])

    lvl = course_level.strip()
    if lvl and not lvl.lower().startswith("level"):
        lvl = f"Level {lvl}"
    course_full = f"{course_name}  {lvl}".strip() if lvl else course_name

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width =  5.33
    ws.column_dimensions["B"].width = 29.55
    ws.column_dimensions["C"].width = 29.22
    ws.column_dimensions["D"].width = 24.55
    ws.column_dimensions["E"].width = 25.55
    ws.column_dimensions["F"].width = 24.00

    # ── Rows 1-3 : Logo (fixed CDACC logo) ────────────────────────────────────
    for r in (1, 2, 3):
        ws.row_dimensions[r].height = 30.0
    ws.merge_cells("A1:F3")

    logo = logo_path or _CDACC_LOGO
    if logo and os.path.exists(logo):
        xl_img  = XLImage(logo)
        logo_px = 85
        marker  = AnchorMarker(col=2, colOff=105 * _EMU, row=0, rowOff=6 * _EMU)
        size    = XDRPositiveSize2D(cx=logo_px * _EMU, cy=logo_px * _EMU)
        xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
        ws.add_image(xl_img)
    _outer_border(ws, 1, 1, 3, 6)

    # ── Row 4 : Organisation name ─────────────────────────────────────────────
    ws.row_dimensions[4].height = 22.0
    ws.merge_cells("A4:F4")
    _set(ws, "A4",
         value="TVET CURRICULUM DEVELOPMENT, ASSESSMENT AND CERTIFICATION COUNCIL (TVET CDACC)",
         font=_f(bold=True, size=12), align=_CTR,
         border=Border(left=_THIN, right=_THIN, top=_NONE, bottom=_THIN))

    # ── Row 5 : Document title ────────────────────────────────────────────────
    ws.row_dimensions[5].height = 22.0
    ws.merge_cells("A5:F5")
    _set(ws, "A5",
         value="SUMMATIVE ASSESSMENT MODERATED PRACTICAL MARKS SHEETS PER UNIT OF COMPETENCY",
         font=_f(bold=True, size=11, color=_BLUE), align=_CTR,
         border=Border(left=_THIN, right=_THIN, top=_NONE, bottom=_THIN))

    # ── Row 6 : Reference code ────────────────────────────────────────────────
    ws.row_dimensions[6].height = 16.0
    ws.merge_cells("A6:F6")
    _set(ws, "A6", value="CBA/SAMSP/1",
         font=_f(bold=True, size=11), align=_LFT,
         border=Border(left=_THIN, right=_THIN, top=_NONE, bottom=_THIN))

    # ── Row 7 : Course/Qualification Code | Title ─────────────────────────────
    ws.row_dimensions[7].height = 25.0
    ws.merge_cells("A7:C7")
    ws.merge_cells("D7:F7")
    ws["A7"].value     = _rich((True,  "Course/Qualification Code:  "),
                               (False, _dots(40)))
    ws["A7"].alignment = _LFT
    ws["D7"].value     = _rich((True,  "Course/Qualification Title:  "),
                               (False, course_full or _dots(40)))
    ws["D7"].alignment = _LFT
    _full_border(ws, 7, 1, 7, 3)
    _full_border(ws, 7, 4, 7, 6)

    # ── Row 8 : Assessment Center Code | Name ─────────────────────────────────
    ws.row_dimensions[8].height = 20.0
    ws.merge_cells("A8:C8")
    ws.merge_cells("D8:F8")
    ws["A8"].value     = _rich((True,  "Assessment Center Code:  "),
                               (False, centre_code or _dots(38)))
    ws["A8"].alignment = _LFT
    ws["D8"].value     = _rich((True,  "Assessment Center Name:  "),
                               (False, centre_name or _dots(38)))
    ws["D8"].alignment = _LFT
    _full_border(ws, 8, 1, 8, 3)
    _full_border(ws, 8, 4, 8, 6)

    # ── Row 9 : Unit Code | Unit Title ────────────────────────────────────────
    ws.row_dimensions[9].height = 20.0
    ws.merge_cells("A9:C9")
    ws.merge_cells("D9:F9")
    ws["A9"].value     = _rich((True,  "Unit Code:  "),
                               (False, unit_code or _dots(45)))
    ws["A9"].alignment = _LFT
    ws["D9"].value     = _rich((True,  "Unit Title:  "),
                               (False, unit_name or _dots(45)))
    ws["D9"].alignment = _LFT
    _full_border(ws, 9, 1, 9, 3)
    _full_border(ws, 9, 4, 9, 6)

    # ── Row 10 : Date of Assessment | Assessment Series ───────────────────────
    ws.row_dimensions[10].height = 20.0
    ws.merge_cells("A10:C10")
    ws.merge_cells("D10:F10")
    ws["A10"].value     = _rich((True,  "Date of Assessment:  From "),
                                (False, _dots(20)),
                                (True,  "  to  "),
                                (False, _dots(20)))
    ws["A10"].alignment = _LFT
    ws["D10"].value     = _rich((True,  "Assessment Series:  "),
                                (False, series or _dots(35)))
    ws["D10"].alignment = _LFT
    _full_border(ws, 10, 1, 10, 3)
    _full_border(ws, 10, 4, 10, 6)

    # ── Row 11 : Column headers ───────────────────────────────────────────────
    ws.row_dimensions[11].height = 18.0
    for coord, label in (
        ("A11", "S/N"),
        ("B11", "Candidate's\nRegistration Code"),
        ("C11", "Candidate's Name"),
        ("D11", "Internal Assessor's Marks\n(All Candidates) (100%)"),
        ("E11", "External Verifier's\n(Sampled Candidates) (100%)"),
        ("F11", "Moderated Marks\n(100%)"),
    ):
        _set(ws, coord, value=label,
             font=_f(bold=True, size=11), align=_CTR,
             fill=_COL_FILL, border=_BORDER_FULL)

    # ── Data rows from row 12 ─────────────────────────────────────────────────
    for i, cand in enumerate(candidates):
        r = 12 + i
        ws.row_dimensions[r].height = 18.0

        ws.cell(row=r, column=1, value=cand["sn"]).alignment     = _CTR
        ws.cell(row=r, column=2, value=cand["reg_no"]).alignment = _LFT
        ws.cell(row=r, column=3, value=cand["name"]).alignment   = _LFT
        ws.cell(row=r, column=4).protection = _UNLOCKED
        ws.cell(row=r, column=5).protection = _UNLOCKED
        ws.cell(row=r, column=6).protection = _UNLOCKED

        for c in range(1, 7):
            ws.cell(row=r, column=c).font   = _f(size=11)
            ws.cell(row=r, column=c).border = _BORDER_FULL
        for c in (1, 4, 5, 6):
            ws.cell(row=r, column=c).alignment = _CTR

    last = 11 + len(candidates)   # last candidate row

    # ── Mean Deviation row ────────────────────────────────────────────────────
    mean_row = last + 2
    ws.row_dimensions[mean_row].height = 20.0
    ws.merge_cells(f"A{mean_row}:B{mean_row}")
    ws.merge_cells(f"C{mean_row}:F{mean_row}")
    ws[f"A{mean_row}"].value     = _rich((True, "Mean Deviation:  "), (False, _dots(18)))
    ws[f"A{mean_row}"].alignment = _LFT
    ws[f"C{mean_row}"].protection = _UNLOCKED
    _full_border(ws, mean_row, 1, mean_row, 6)

    # ── Footer / Signature rows ───────────────────────────────────────────────
    # Each signatory gets: (1) full-width name row  (2) split ID | Sig | Date row
    ia_name = mean_row + 2
    ia_sig  = ia_name + 1
    ev_name = ia_sig  + 1
    ev_sig  = ev_name + 1
    cm_name = ev_sig  + 1
    cm_sig  = cm_name + 1

    # Name rows — full width
    for r, label in (
        (ia_name, "1. Name of Internal Assessor:  "),
        (ev_name, "2. Name of External Verifier:  "),
        (cm_name, "3. Name of Assessment Center Manager/Officer:  "),
    ):
        ws.row_dimensions[r].height = 20.0
        ws.merge_cells(f"A{r}:F{r}")
        ws[f"A{r}"].value     = _rich((True, label), (False, _dots(50)))
        ws[f"A{r}"].alignment = _LFT
        _full_border(ws, r, 1, r, 6)

    # Internal Assessor & External Verifier: A:D | E | F
    for r, label in (
        (ia_sig, "Internal Assessor Reg. Code/National ID. No.: "),
        (ev_sig, "External Verifier Reg. Code/National ID. No.: "),
    ):
        ws.row_dimensions[r].height = 22.0
        ws.merge_cells(f"A{r}:D{r}")
        ws[f"A{r}"].value     = _rich((True, label), (False, _dots(20)))
        ws[f"A{r}"].alignment = _LFT
        ws[f"E{r}"].value     = _rich((True, "Signature:  "), (False, _dots(20)))
        ws[f"E{r}"].alignment = _LFT
        ws[f"F{r}"].value     = _rich((True, "Date:  "), (False, _dots(15)))
        ws[f"F{r}"].alignment = _LFT
        _full_border(ws, r, 1, r, 6)

    # Centre Manager last row: A:C | D | E:F
    ws.row_dimensions[cm_sig].height = 22.0
    ws.merge_cells(f"A{cm_sig}:C{cm_sig}")
    ws.merge_cells(f"E{cm_sig}:F{cm_sig}")
    ws[f"A{cm_sig}"].value     = _rich((True, "Signature: "),           (False, _dots(20)))
    ws[f"A{cm_sig}"].alignment = _LFT
    ws[f"D{cm_sig}"].value     = _rich((True, "Date:  "),               (False, _dots(15)))
    ws[f"D{cm_sig}"].alignment = _LFT
    ws[f"E{cm_sig}"].value     = _rich((True, "Institutional Stamp: "), (False, _dots(20)))
    ws[f"E{cm_sig}"].alignment = _LFT
    _full_border(ws, cm_sig, 1, cm_sig, 6)

    # ── Print settings ────────────────────────────────────────────────────────
    ws.print_area = f"A1:F{cm_sig}"
    ws.page_setup.paperSize   = 9           # A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left   = 0.5
    ws.page_margins.right  = 0.5
    ws.page_margins.top    = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.3
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.center.font = _FONT
    ws.oddFooter.center.size = 11
    ws.print_title_rows = "11:11"
    ws.protection.sheet               = True
    ws.protection.password            = "2026"
    ws.protection.selectLockedCells   = False
    ws.protection.selectUnlockedCells = False


# ── Public entry point ────────────────────────────────────────────────────────

def build_summative_per_unit(data: dict, logo_path: str = None) -> list[tuple[str, bytes]]:
    """One summative moderated-practical marks sheet per unit.
    Returns [(filename, file_bytes), ...]."""
    results: list[tuple[str, bytes]] = []
    for unit in data["units"]:
        wb = Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet(
            _safe_sheet_name(
                re.sub(r"[/\\:*?\[\]]", "-", unit["unit_name"])[:31], set()
            )
        )
        build_summative_sheet(ws, data, unit, logo_path)
        buf = io.BytesIO()
        wb.save(buf)
        fname = _safe_filename(unit["unit_name"], unit.get("unit_code", "")) + ".xlsx"
        results.append((fname, buf.getvalue()))
    return results
