#!/usr/bin/env python3
"""
extract_nominal.py
==================
Parse a TVET CDACC Nominal Roll PDF and return structured data compatible
with marksheet_excel.build_marksheet_per_unit / build_marksheet_workbook.

Each page table is processed row by row.  Header rows carry a recognised key
in column 0 (e.g. "Report Type:", "Unit:", "Assessment Center:").  Candidate
rows carry a digit SN in column 0 and a CDACC reg-no pattern in column 2.

Unique record key: (CDACC reg No., unit_name) — prevents duplicates that arise
from the same candidate appearing on continuation pages that repeat the header.

Usage (CLI):
    python extract_nominal.py nominal_roll.pdf

Requires: pdfplumber
"""

import argparse
import json
import re
import sys
from collections import OrderedDict

import pdfplumber

# ── Patterns ──────────────────────────────────────────────────────────────────

TITLE_RE = re.compile(r'NOMINAL ROLL FOR:\s*\(([^)]+)\)', re.I)
CDACC_RE = re.compile(r'^[A-Z0-9]{6,9}/[A-Z]+/\d+/\d{4}/\d+$')


# ── Helpers ───────────────────────────────────────────────────────────────────

def _clean(s):
    return re.sub(r'\s+', ' ', (s or '').strip())


def _parse_centre(val):
    """'RIFT VALLEY INSTITUTE OF SCIENCE AND TECHNOLOGY(RVIST) ( 0320134P )'
    → ('RIFT VALLEY INSTITUTE OF SCIENCE AND TECHNOLOGY(RVIST)', '0320134P')
    """
    m = re.search(r'\(\s*([A-Z0-9]+)\s*\)\s*$', val)
    if m:
        return val[:m.start()].strip(), m.group(1)
    return val.strip(), ''


# ── Main extractor ────────────────────────────────────────────────────────────

def extract(pdf_path, on_log=None, on_progress=None):
    """Parse the nominal roll PDF and return a data dict.

    The returned dict mirrors the schema expected by marksheet_excel:
      {
        centre_name, centre_code, course_name, course_level, series,
        unit_count,
        units: [
          { unit_code, unit_name, report_type, course_name, course_level,
            candidate_count,
            candidates: [{ sn, reg_no, admission_no, name }, ...] }
        ]
      }
    """
    def _log(level, msg):
        if on_log:
            on_log(level, msg)

    state = {
        'series':       None,
        'report_type':  None,
        'centre_name':  None,
        'centre_code':  None,
        'course_level': None,
        'course_name':  None,
    }
    # Key: (report_type, unit_name) — exact character-by-character match.
    # "Control ICT Security Threats" != "Control Ict Security Threats" → separate.
    # "Control ICT Security Threats" == "Control ICT Security Threats" → merged.
    units        = OrderedDict()   # (report_type, unit_name) → unit dict
    current_unit = None

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        _log('step', f'Opened — {total} page(s) detected')

        for pg_idx, page in enumerate(pdf.pages, 1):
            _log('info', f'Reading page {pg_idx} / {total}')
            if on_progress:
                on_progress(pg_idx, total)

            text = page.extract_text() or ''

            # Exam series from title line (first occurrence)
            if state['series'] is None:
                m = TITLE_RE.search(text)
                if m:
                    state['series'] = m.group(1).strip()
                    _log('found', f'Series: {state["series"]}')

            for table in (page.extract_tables() or []):
                for row in table:
                    if not row:
                        continue

                    # Normalise key: strip whitespace / newlines / trailing colon
                    key = _clean(row[0]).rstrip(':')
                    val = _clean(row[1]) if len(row) > 1 and row[1] else ''

                    # ── Header metadata rows ──────────────────────────────────

                    if 'Report Type' in key:
                        state['report_type'] = val
                        continue

                    if 'Assessment' in key and 'Center' in key:
                        name, code = _parse_centre(val)
                        state['centre_name'] = name
                        state['centre_code'] = code
                        continue

                    if 'Course Level' in key:
                        state['course_level'] = val
                        continue

                    if key == 'Course':
                        state['course_name'] = val
                        continue

                    if key == 'Unit':
                        unit_name = val
                        rt        = (state['report_type'] or '').strip()
                        ukey      = (rt, unit_name)   # exact match — no lowercasing
                        if ukey not in units:
                            units[ukey] = {
                                'unit_code':    '',
                                'unit_name':    unit_name,
                                'report_type':  rt,
                                'course_name':  state['course_name'],
                                'course_level': state['course_level'],
                                'candidates':   [],
                                '_seen':        set(),   # CDACC reg nos in this unit
                            }
                            _log('found', f'Unit: {unit_name}  [{rt}]')
                        current_unit = units[ukey]
                        continue

                    # Skip column-header and decorative rows
                    if key in ('Date', 'SN', 'Candidate Photo',
                               'Candidate\nPhoto', 'CDACC reg No'):
                        continue

                    # ── Candidate data row ────────────────────────────────────

                    if current_unit is None or len(row) < 5:
                        continue

                    sn_raw = _clean(row[0])
                    cdacc  = re.sub(r'\s+', '', row[2] or '')   # no spaces
                    adm    = re.sub(r'\s+', '', row[3] or '')
                    name   = _clean(row[4])

                    if sn_raw.isdigit() and CDACC_RE.match(cdacc):
                        if cdacc not in current_unit['_seen']:
                            current_unit['_seen'].add(cdacc)
                            current_unit['candidates'].append({
                                'sn':           int(sn_raw),
                                'reg_no':       cdacc,
                                'admission_no': adm,
                                'name':         name,
                            })

    # ── Post-process ──────────────────────────────────────────────────────────
    result_units = []
    for u in units.values():
        u.pop('_seen', None)
        # Re-number SN sequentially within each merged unit
        for i, c in enumerate(u['candidates'], 1):
            c['sn'] = i
        u['candidate_count'] = len(u['candidates'])
        result_units.append(u)

    _log('success',
         f'Done — {len(result_units)} unit(s), '
         f'{sum(u["candidate_count"] for u in result_units)} registrations')

    return {
        'centre_name':  state['centre_name'],
        'centre_code':  state['centre_code'],
        'course_name':  state['course_name'],
        'course_level': state['course_level'],
        'series':       state['series'],
        'unit_count':   len(result_units),
        'units':        result_units,
    }


# ── CLI ───────────────────────────────────────────────────────────────────────

def _print_summary(data):
    print(f"Centre : {data['centre_name']} ({data['centre_code']})")
    print(f"Course : {data['course_name']}  {data['course_level']}")
    print(f"Series : {data['series']}")
    print(f"Units  : {data['unit_count']}\n")
    print(f"{'#':>4}  {'UNIT NAME':<45}  {'TYPE'}")
    print("-" * 72)
    for u in data['units']:
        rt = '[RE-ASSESS]' if (u['report_type'] or '').lower().startswith('re') else ''
        print(f"{u['candidate_count']:>4}  {u['unit_name']:<50}  {rt}")
    print("-" * 72)
    total = sum(u['candidate_count'] for u in data['units'])
    print(f"{total:>4}  TOTAL")


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Extract TVET CDACC Nominal Roll data from a PDF.")
    p.add_argument("pdf", help="path to the nominal roll PDF")
    p.add_argument("--json", metavar="FILE",
                   help="write full structured data to JSON")
    args = p.parse_args(argv)

    data = extract(args.pdf)

    if args.json:
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"Wrote JSON → {args.json}", file=sys.stderr)

    _print_summary(data)
    return data


if __name__ == '__main__':
    main()
