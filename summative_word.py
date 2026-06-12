#!/usr/bin/env python3
"""
summative_word.py
=================
Word (.docx) rendition of the TVET CDACC Summative Assessment Moderated
Practical Marks Sheet — one document per unit of competency.

Formatting follows "Summative Word Sample.docx" exactly.  The styles, page
section AND the signatory block come verbatim from summative_word_template.docx
(a stripped copy of that sample that keeps only the signatory table), so the
signatory block is byte-for-byte the user's own.

Behaviour:
  • A4 landscape; logo, org (13pt), CBA/SAMSP/1, blue title (12pt).
  • Borderless 2-col info grid (Plain Table 4, no row shading).
  • Grey-headed candidate table (Table Grid); its header row repeats on every
    page while there are candidate rows.
  • The "Mean Deviation" line and the signatory block are kept together on the
    same page.
  • The info grid and the signatory block are locked against editing (content
    controls); only the candidate list stays editable.
"""

import io
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
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

# Course/Qualification Code, Unit Code and Date of Assessment are reproduced
# verbatim from the sample: blank, filled with the ellipsis (…) leader it uses.
_ELL          = "…"
_COURSE_CODE  = "Course/Qualification Code: " + _ELL * 22 + "."
_UNIT_CODE    = "Unit Code: " + _ELL * 29 + "."
_DATE_ASSESS  = "Date of Assessment:  From " + _ELL * 10 + "." + "  to  " + _ELL * 10

# Shown as the locked content control's label when a user clicks the area.
_LOCK_MSG = "Locked — you can add this data by hand after printing."


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


def _no_band(table):
    """Disable the Plain Table 4 row banding so the table shows no shading
    (matches the sample's tblLook: noHBand=1, noVBand=1)."""
    tbl_pr = table._tbl.tblPr
    look = tbl_pr.find(qn("w:tblLook"))
    if look is None:
        look = OxmlElement("w:tblLook")
        tbl_pr.append(look)
    for attr, val in (("val", "06A0"), ("firstRow", "1"), ("lastRow", "0"),
                      ("firstColumn", "1"), ("lastColumn", "0"),
                      ("noHBand", "1"), ("noVBand", "1")):
        look.set(qn("w:" + attr), val)


def _repeat_header(table):
    """Mark row 0 as a table header row so it repeats on every page."""
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        th = OxmlElement("w:tblHeader")
        th.set(qn("w:val"), "true")
        tr_pr.append(th)


def _keep_next(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    if pPr.find(qn("w:keepNext")) is None:
        pPr.append(OxmlElement("w:keepNext"))


def _keep_table_together(table):
    """No row splits across a page, and every row keeps with the next, so the
    whole table stays on one page."""
    rows = table.rows
    for i, row in enumerate(rows):
        tr_pr = row._tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))
        if i < len(rows) - 1:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _keep_next(p)


def _lock_table(table, sdt_id: int, message: str):
    """Wrap a table in a content control locked against content edits.  The
    message becomes the control's title/tag, shown as a label when the user
    clicks the locked area (so they're told they can fill it in after print)."""
    tbl = table._tbl
    parent = tbl.getparent()
    idx = parent.index(tbl)
    parent.remove(tbl)

    sdt = OxmlElement("w:sdt")
    sdt_pr = OxmlElement("w:sdtPr")
    # sdtPr child order: alias, tag, id, lock
    alias = OxmlElement("w:alias"); alias.set(qn("w:val"), message); sdt_pr.append(alias)
    tag = OxmlElement("w:tag"); tag.set(qn("w:val"), message); sdt_pr.append(tag)
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

    # The signatory block is the template's only table (verbatim from sample).
    sig = doc.tables[0]

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

    # ── Info grid (Plain Table 4, no shading, full width) ─────────────────────
    # Left column: Code / Code / Code / Date — the latter three taken verbatim
    # from the sample (blank, ellipsis leader). Right column: filled from data.
    info = doc.add_table(rows=4, cols=2, style="Plain Table 4")
    left_cells = [
        [(True, "Assessment Center Code:  "), (False, centre_code or _dots())],
        [(False, _COURSE_CODE)],
        [(False, _UNIT_CODE)],
        [(False, _DATE_ASSESS)],
    ]
    right_cells = [
        [(True, "Assessment Center Name:  "),     (False, centre_name or _dots())],
        [(True, "Course/Qualification Title:  "), (False, course_full or _dots())],
        [(True, "Unit Title:  "),                 (False, unit_name or _dots())],
        [(True, "Assessment Series:  "),          (False, series or _dots())],
    ]
    for r in range(4):
        _cell(info.rows[r].cells[0], left_cells[r])
        _cell(info.rows[r].cells[1], right_cells[r])
        _row_height(info.rows[r], 0.43)
    _fix_widths(info, [7699, 7699])
    _no_band(info)

    # ── Candidate table (Table Grid, grey header that repeats per page) ───────
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
    _repeat_header(table)

    # ── Mean Deviation line (kept together with the signatory block) ──────────
    for _ in range(3):
        doc.add_paragraph()
    keep = [_para(doc, [(True, "Mean Deviation:  ")], size=11)]
    keep += [doc.add_paragraph() for _ in range(3)]

    # ── Place the signatory block here (verbatim from the template) ───────────
    trailing = doc.add_paragraph()
    trailing._p.addprevious(sig._tbl)

    for p in keep:
        _keep_next(p)
    _keep_table_together(sig)

    # ── Lock the two side tables; the candidate list stays editable ───────────
    _lock_table(info, 101, _LOCK_MSG)
    _lock_table(sig, 102, _LOCK_MSG)


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
