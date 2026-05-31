#!/usr/bin/env python3
"""
summative_sample.py
===================
Generate a ONE-UNIT sample of the Summative Assessment Moderated Practical
Marks Sheet so the layout can be reviewed before building the full generator.

Run:  python summative_sample.py
Output: summative_sample.xlsx
"""

import os

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, InlineFont, TextBlock
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import (
    AnchorMarker, OneCellAnchor, XDRPositiveSize2D,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side

# ── Paths ─────────────────────────────────────────────────────────────────────

_HERE      = os.path.dirname(os.path.abspath(__file__))
_LOGO_PATH = os.path.join(_HERE, "cdacc_logo.png")
_EMU       = 9525

# ── Style constants ───────────────────────────────────────────────────────────

_FONT   = "Times New Roman"
_BLUE   = "1F4E79"
_GREY   = "D9D9D9"

_NONE = Side(style=None)
_THIN = Side(style="thin", color="000000")
_MED  = Side(style="medium", color="000000")

_BORDER_FULL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HDR_FILL    = PatternFill("solid", fgColor="EBF3FB")   # light-blue header tint
_COL_FILL    = PatternFill("solid", fgColor=_GREY)
_UNLOCKED    = Protection(locked=False)

_CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LFT  = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_LFT_T = Alignment(horizontal="left",  vertical="top",    wrap_text=True)


def _f(bold=False, size=11, color="000000"):
    return Font(name=_FONT, bold=bold, size=size, color=color)


def _rich(*parts, size=11):
    """Build CellRichText from alternating (bold, text) pairs."""
    blocks = []
    for bold, text in parts:
        blocks.append(TextBlock(InlineFont(b=bold, rFont=_FONT, sz=size), text))
    return CellRichText(*blocks)


def _set(ws, coord, value=None, font=None, align=None, fill=None, border=None):
    c = ws[coord]
    if value  is not None: c.value     = value
    if font   is not None: c.font      = font
    if align  is not None: c.alignment = align
    if fill   is not None: c.fill      = fill
    if border is not None: c.border    = border


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


def _full_border(ws, r1, c1, r2, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = _BORDER_FULL


def _dots(n=50):
    return "." * n


# ── Sheet builder ─────────────────────────────────────────────────────────────

def build_summative_sheet(ws, data: dict, unit: dict):
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

    # ── Rows 1-3 : Logo ───────────────────────────────────────────────────────
    for r in (1, 2, 3):
        ws.row_dimensions[r].height = 30.0
    ws.merge_cells("A1:F3")

    if os.path.exists(_LOGO_PATH):
        xl_img  = XLImage(_LOGO_PATH)
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
         font=_f(bold=True, size=13), align=_CTR,
         border=Border(left=_THIN, right=_THIN, top=_NONE, bottom=_THIN))

    # ── Row 5 : Document title ────────────────────────────────────────────────
    ws.row_dimensions[5].height = 22.0
    ws.merge_cells("A5:F5")
    _set(ws, "A5",
         value="SUMMATIVE ASSESSMENT MODERATED PRACTICAL MARKS SHEETS PER UNIT OF COMPETENCY",
         font=_f(bold=True, size=12, color=_BLUE), align=_CTR,
         border=Border(left=_THIN, right=_THIN, top=_NONE, bottom=_THIN))

    # ── Row 6 : Reference code ────────────────────────────────────────────────
    ws.row_dimensions[6].height = 16.0
    ws.merge_cells("A6:F6")
    _set(ws, "A6", value="CBA/SAMSP/1",
         font=_f(bold=True, size=11), align=_LFT,
         border=Border(left=_THIN, right=_THIN, top=_NONE, bottom=_THIN))

    # ── Row 7 : Course/Qualification Code | Course/Qualification Title ────────
    ws.row_dimensions[7].height = 20.0
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

    # ── Row 11 : Assessor 1 | Mean Deviation ─────────────────────────────────
    ws.row_dimensions[11].height = 20.0
    ws.merge_cells("A11:D11")
    ws.merge_cells("E11:F11")
    ws["A11"].value     = _rich((True,  "1. Name of Internal Assessor:  "),
                                (False, _dots(50)))
    ws["A11"].alignment = _LFT
    ws["E11"].value     = _rich((True,  "Mean Deviation:  "),
                                (False, _dots(18)))
    ws["E11"].alignment = _LFT
    _full_border(ws, 11, 1, 11, 4)
    _full_border(ws, 11, 5, 11, 6)

    # ── Row 12 : Assessor 2 ───────────────────────────────────────────────────
    ws.row_dimensions[12].height = 20.0
    ws.merge_cells("A12:F12")
    ws["A12"].value     = _rich((True,  "2. Name of External Verifier:  "),
                                (False, _dots(80)))
    ws["A12"].alignment = _LFT
    _full_border(ws, 12, 1, 12, 6)

    # ── Row 13 : Assessment Centre Manager ────────────────────────────────────
    ws.row_dimensions[13].height = 20.0
    ws.merge_cells("A13:F13")
    ws["A13"].value     = _rich((True,  "3. Name of Assessment Center Manager/Officer:  "),
                                (False, _dots(60)))
    ws["A13"].alignment = _LFT
    _full_border(ws, 13, 1, 13, 6)

    # ── Row 14 : Column headers ───────────────────────────────────────────────
    ws.row_dimensions[14].height = 42.0
    for coord, label in (
        ("A14", "S/N"),
        ("B14", "Candidate's\nRegistration Code"),
        ("C14", "Candidate's Name"),
        ("D14", "Internal Assessor's Marks\n(All Candidates) (100%)"),
        ("E14", "External Verifier's\n(Sampled Candidates) (100%)"),
        ("F14", "Moderated Marks\n(100%)"),
    ):
        _set(ws, coord, value=label,
             font=_f(bold=True, size=11), align=_CTR,
             fill=_COL_FILL, border=_BORDER_FULL)

    # ── Data rows from row 15 ─────────────────────────────────────────────────
    for i, cand in enumerate(candidates):
        r = 15 + i
        ws.row_dimensions[r].height = 18.0

        ws.cell(row=r, column=1, value=cand["sn"]).alignment          = _CTR
        ws.cell(row=r, column=2, value=cand["reg_no"]).alignment      = _LFT
        ws.cell(row=r, column=3, value=cand["name"]).alignment        = _LFT
        # Columns D, E — editable (marks)
        ws.cell(row=r, column=4).protection = _UNLOCKED
        ws.cell(row=r, column=5).protection = _UNLOCKED
        # Column F — moderated marks (editable)
        ws.cell(row=r, column=6).protection = _UNLOCKED

        for c in range(1, 7):
            ws.cell(row=r, column=c).font   = _f(size=11)
            ws.cell(row=r, column=c).border = _BORDER_FULL
        for c in (1, 4, 5, 6):
            ws.cell(row=r, column=c).alignment = _CTR

    # ── Footer / Signature rows ───────────────────────────────────────────────
    last         = 14 + len(candidates)
    sig1_row     = last + 2
    sig2_row     = last + 3
    sig3_row     = last + 4

    for r in (last + 1, sig1_row, sig2_row, sig3_row):
        ws.row_dimensions[r].height = 22.0

    ws.merge_cells(f"A{sig1_row}:F{sig1_row}")
    ws[f"A{sig1_row}"].value = _rich(
        (True,  "Internal Assessor Reg. Code/National ID. No.: "),
        (False, _dots(20)),
        (True,  "  Signature: "),
        (False, _dots(20)),
        (True,  "  Date: "),
        (False, _dots(15)),
    )
    ws[f"A{sig1_row}"].alignment = _LFT
    _full_border(ws, sig1_row, 1, sig1_row, 6)

    ws.merge_cells(f"A{sig2_row}:F{sig2_row}")
    ws[f"A{sig2_row}"].value = _rich(
        (True,  "External Verifier Reg. Code/National ID. No.: "),
        (False, _dots(20)),
        (True,  "  Signature: "),
        (False, _dots(20)),
        (True,  "  Date: "),
        (False, _dots(15)),
    )
    ws[f"A{sig2_row}"].alignment = _LFT
    _full_border(ws, sig2_row, 1, sig2_row, 6)

    ws.merge_cells(f"A{sig3_row}:F{sig3_row}")
    ws[f"A{sig3_row}"].value = _rich(
        (True,  "Signature: "),
        (False, _dots(25)),
        (True,  "  Date: "),
        (False, _dots(20)),
        (True,  "  Institutional Stamp: "),
        (False, _dots(25)),
    )
    ws[f"A{sig3_row}"].alignment = _LFT
    _full_border(ws, sig3_row, 1, sig3_row, 6)

    # ── Print settings ────────────────────────────────────────────────────────
    ws.print_area = f"A1:F{sig3_row}"
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
    ws.print_title_rows = "14:14"
    ws.protection.sheet               = True
    ws.protection.password            = "2026"
    ws.protection.selectLockedCells   = False
    ws.protection.selectUnlockedCells = False


# ── Sample data ───────────────────────────────────────────────────────────────

SAMPLE_DATA = {
    "centre_code":  "0320134P",
    "centre_name":  "RIFT VALLEY NATIONAL POLYTECHNIC",
    "course_name":  "INFORMATION COMMUNICATION TECHNOLOGY",
    "course_level": "Level 5",
    "series":       "November/December 2025",
    "units": [
        {
            "unit_code":    "ICT/OS/CU/CR/01/5/A",
            "unit_name":    "Configure ICT Security Threats",
            "report_type":  "Assessment",
            "course_name":  "INFORMATION COMMUNICATION TECHNOLOGY",
            "course_level": "Level 5",
            "candidate_count": 8,
            "candidates": [
                {"sn": 1, "reg_no": "RVIST/ICT/CR/001/2025/001", "admission_no": "ADM001", "name": "ACHIENG JOYCE ATIENO"},
                {"sn": 2, "reg_no": "RVIST/ICT/CR/001/2025/002", "admission_no": "ADM002", "name": "BETT KIPRONO JAMES"},
                {"sn": 3, "reg_no": "RVIST/ICT/CR/001/2025/003", "admission_no": "ADM003", "name": "CHEBET FAITH JELIMO"},
                {"sn": 4, "reg_no": "RVIST/ICT/CR/001/2025/004", "admission_no": "ADM004", "name": "GITAU DAVID MWANGI"},
                {"sn": 5, "reg_no": "RVIST/ICT/CR/001/2025/005", "admission_no": "ADM005", "name": "HASSAN IBRAHIM OMAR"},
                {"sn": 6, "reg_no": "RVIST/ICT/CR/001/2025/006", "admission_no": "ADM006", "name": "KAMAU WANJIKU GRACE"},
                {"sn": 7, "reg_no": "RVIST/ICT/CR/001/2025/007", "admission_no": "ADM007", "name": "LANGAT KIPCHOGE ENOCH"},
                {"sn": 8, "reg_no": "RVIST/ICT/CR/001/2025/008", "admission_no": "ADM008", "name": "MUTUA STEPHEN NDULI"},
            ],
        }
    ],
}


if __name__ == "__main__":
    wb = Workbook()
    ws = wb.active
    ws.title = "Configure ICT Security"

    build_summative_sheet(ws, SAMPLE_DATA, SAMPLE_DATA["units"][0])

    out = "summative_sample.xlsx"
    wb.save(out)
    print(f"Saved: {out}")
