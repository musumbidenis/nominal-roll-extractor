#!/usr/bin/env python3
"""
summative_word.py
=================
Word (.docx) rendition of the TVET CDACC Summative Assessment Moderated
Practical Marks Sheet — one document per unit of competency.

Formatting follows "Summative Word Sample.docx" exactly: A4 landscape, the
logo + headings, a borderless 2-column info grid (Plain Table 4), the
grey-headed candidate table (Table Grid), a "Mean Deviation" line, and a
borderless signatory block (Plain Table 4).  The styles and page section are
inherited from summative_word_template.docx (a stripped copy of that sample),
so the look is identical.

The two side tables — the info grid and the signatory block — are wrapped in
content controls locked against editing, so only the candidate list (where
marks are entered) can be changed.
"""

import io
import os

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Twips

from marksheet_excel import _safe_filename

_HERE       = os.path.dirname(os.path.abspath(__file__))
_TEMPLATE   = os.path.join(_HERE, "summative_word_template.docx")
_CDACC_LOGO = os.path.join(_HERE, "cdacc_logo.png")
_FONT       = "Times New Roman"
_BLUE       = RGBColor(0x1F, 0x4E, 0x79)
_GREY       = "D9D9D9"

_ORG   = "TVET CURRICULUM DEVELOPMENT, ASSESSMENT AND CERTIFICATION COUNCIL (TVET CDACC)"
_TITLE = "SUMMATIVE ASSESSMENT MODERATED PRACTICAL MARKS SHEETS PER UNIT OF COMPETENCY"

_HEADERS = ["S/N", "Candidate's\nRegistration Code", "Candidate's Name",
            "Internal Assessor's Marks\n(All Candidates) (100%)",
            "External Verifier's\n(Sampled Candidates) (100%)",
            "Moderated Marks\n(100%)"]
# candidate-table column widths, in twips, taken verbatim from the sample
_COL_TW  = [920, 3678, 3173, 2662, 2662, 2293]

# signatory block — exact lines from the sample (all bold)
_SIG_LINES = [
    "1. Name of Internal Assessor:  " + "." * 200,
    "Internal Assessor Reg. Code/National ID. No.: " + "." * 73 +
        "\tSignature: " + "." * 25 + "  Date: " + "." * 35,
    "2. Name of External Verifier:  " + "." * 200,
    "External Verifier Reg. Code/National ID. No.: " + "." * 73 +
        "\tSignature: " + "." * 25 + "  Date: " + "." * 35,
    "3. Name of Assessment Center Manager/Officer:  " + "." * 185,
    "Signature: " + "." * 25 + "  Date: " + "." * 35 +
        " Institutional Stamp: " + "." * 110,
]


def _dots(n: int = 28) -> str:
    return "." * n


# ── Low-level helpers ─────────────────────────────────────────────────────────

def _style_run(run, size=11, bold=False, color=None):
    run.bold = bool(bold)
    run.font.name = _FONT
    run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color


def _para(doc, parts, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, color=None,
          space_after=2):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(space_after)
    for bold, text in parts:
        _style_run(p.add_run(text), size=size, bold=bold, color=color)
    return p


def _cell(cell, parts, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, shade=None):
    p = cell.paragraphs[0]
    p.text = ""
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    for bold, text in parts:
        chunks = str(text).split("\n")
        for j, chunk in enumerate(chunks):
            _style_run(p.add_run(chunk), size=size, bold=bold)
            if j < len(chunks) - 1:
                p.runs[-1].add_break()
    cell.vertical_alignment = 1  # centre vertically
    if shade:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), shade)
        cell._tc.get_or_add_tcPr().append(shd)


def _row_height(row, inches: float):
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    row.height = Inches(inches)


def _fix_widths(table, twips_list):
    table.autofit = False
    table.allow_autofit = False
    for row in table.rows:
        for ci, tw in enumerate(twips_list):
            row.cells[ci].width = Twips(tw)


def _lock_table(table, sdt_id: int):
    """Wrap a table in a content control locked against content edits, so the
    table can't be modified while the rest of the document stays editable."""
    tbl = table._tbl
    parent = tbl.getparent()
    idx = parent.index(tbl)
    parent.remove(tbl)

    sdt = OxmlElement("w:sdt")
    sdt_pr = OxmlElement("w:sdtPr")
    _id = OxmlElement("w:id"); _id.set(qn("w:val"), str(sdt_id)); sdt_pr.append(_id)
    lock = OxmlElement("w:lock"); lock.set(qn("w:val"), "sdtContentLocked")
    sdt_pr.append(lock)
    sdt.append(sdt_pr)
    sdt_content = OxmlElement("w:sdtContent")
    sdt_content.append(tbl)
    sdt.append(sdt_content)
    parent.insert(idx, sdt)


# ── Document builder ──────────────────────────────────────────────────────────

def build_summative_doc(doc, data: dict, unit: dict):
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

    # ── Logo (centred) ────────────────────────────────────────────────────────
    if os.path.exists(_CDACC_LOGO):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        p.add_run().add_picture(_CDACC_LOGO, width=Inches(0.96), height=Inches(0.9))

    # ── Headings ──────────────────────────────────────────────────────────────
    _para(doc, [(True, _ORG)],   size=13, align=WD_ALIGN_PARAGRAPH.CENTER)
    _para(doc, [(True, "CBA/SAMSP/1")], size=11)
    _para(doc, [(True, _TITLE)], size=12, align=WD_ALIGN_PARAGRAPH.CENTER, color=_BLUE)
    doc.add_paragraph()

    # ── Info grid (Plain Table 4, borderless, full width) ─────────────────────
    info = doc.add_table(rows=4, cols=2, style="Plain Table 4")
    rows = [
        (("Assessment Center Code:  ",     centre_code or _dots()),
         ("Assessment Center Name:  ",     centre_name or _dots())),
        (("Course/Qualification Code:  ",  _dots()),
         ("Course/Qualification Title:  ", course_full or _dots())),
        (("Unit Code:  ",                  unit_code or _dots()),
         ("Unit Title:  ",                 unit_name or _dots())),
        (("Date of Assessment:  From ",    _dots(12) + ".  to  " + _dots(12)),
         ("Assessment Series:  ",          series or _dots())),
    ]
    for r, (left, right) in enumerate(rows):
        _cell(info.rows[r].cells[0], [(True, left[0]),  (False, left[1])])
        _cell(info.rows[r].cells[1], [(True, right[0]), (False, right[1])])
        _row_height(info.rows[r], 0.43)
    _fix_widths(info, [7699, 7699])

    # ── Candidate table (Table Grid, grey header) ─────────────────────────────
    table = doc.add_table(rows=1 + len(candidates), cols=6, style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for c, label in enumerate(_HEADERS):
        _cell(table.rows[0].cells[c], [(True, label)],
              align=WD_ALIGN_PARAGRAPH.CENTER, shade=_GREY)
    _row_height(table.rows[0], 0.58)

    for i, cand in enumerate(candidates, 1):
        cells = table.rows[i].cells
        _cell(cells[0], [(False, str(cand["sn"]))], align=WD_ALIGN_PARAGRAPH.CENTER)
        _cell(cells[1], [(False, cand["reg_no"])])
        _cell(cells[2], [(False, cand["name"])])
        for c in (3, 4, 5):
            _cell(cells[c], [(False, "")], align=WD_ALIGN_PARAGRAPH.CENTER)
        _row_height(table.rows[i], 0.25)

    _fix_widths(table, _COL_TW)

    # ── Mean Deviation line (with the sample's surrounding spacing) ───────────
    for _ in range(3):
        doc.add_paragraph()
    _para(doc, [(True, "Mean Deviation:  ")], size=11)
    for _ in range(3):
        doc.add_paragraph()

    # ── Signatory block (Plain Table 4, borderless, full width) ───────────────
    sig = doc.add_table(rows=len(_SIG_LINES), cols=1, style="Plain Table 4")
    for r, line in enumerate(_SIG_LINES):
        _cell(sig.rows[r].cells[0], [(True, line)])
    _fix_widths(sig, [15398])

    doc.add_paragraph()   # trailing paragraph (matches the sample)

    # ── Lock the two side tables; leave the candidate list editable ───────────
    _lock_table(info, 101)
    _lock_table(sig, 102)


# ── Public entry point ────────────────────────────────────────────────────────

def build_summative_per_unit_docx(data: dict) -> list[tuple[str, bytes]]:
    """One summative moderated-practical marks sheet (.docx) per unit.
    Returns [(filename, file_bytes), ...]."""
    results: list[tuple[str, bytes]] = []
    for unit in data["units"]:
        doc = Document(_TEMPLATE)
        build_summative_doc(doc, data, unit)
        buf = io.BytesIO()
        doc.save(buf)
        fname = _safe_filename(unit["unit_name"], unit.get("unit_code", "")) + ".docx"
        results.append((fname, buf.getvalue()))
    return results
