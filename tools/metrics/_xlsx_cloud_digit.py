# -*- coding: utf-8 -*-
"""Probe: the digit width Excel measures for a Normal font GDI cannot see
(the Office cloud faces), per size. A 100-digit column is 100 digits wide to
the pixel, so Width/0.75 // 100 is the digit.

    python tools/metrics/_xlsx_cloud_digit.py
"""
import json
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "xlsx_probes" / "cloud_digit.json"
FACES = ["Aptos Narrow", "Aptos", "Aptos Display"]
SIZES = [x / 2 for x in range(12, 145)]  # 6pt .. 72pt by halves


def main():
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    got = {}
    try:
        wb = xl.Workbooks.Add(); sh = wb.Worksheets(1)
        normal = wb.Styles("Normal").Font
        for face in FACES:
            normal.Name = face
            for size in SIZES:
                normal.Size = size
                sh.Columns(1).ColumnWidth = 100
                px = round(sh.Columns(1).Width / 0.75)
                got.setdefault(face, {})[str(size)] = px // 100
        wb.Close(False)
    finally:
        xl.Quit()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(got, indent=1))
    for face, by in got.items():
        print(face, [(s, d) for s, d in list(by.items())[:12]])


if __name__ == "__main__":
    main()
