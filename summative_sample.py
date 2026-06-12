#!/usr/bin/env python3
"""
summative_sample.py
===================
Generate a ONE-UNIT sample of the Summative Assessment Moderated Practical
Marks Sheet so the layout can be eyeballed.  The sheet builder now lives in
summative_excel.py; this script just drives it with sample data.

Run:  python summative_sample.py
Output: summative_sample.xlsx
"""

from openpyxl import Workbook

from summative_excel import build_summative_sheet

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
