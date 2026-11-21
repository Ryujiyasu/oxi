# -*- coding: utf-8 -*-
"""Probe: the height Excel gives a row whose only cell is in font F at size S,
bold or not, with <charset val=C/> (none / 0 / 128 / 134) on the font record.

One workbook, one row per arm, no ht written, so Excel derives every height
itself on open. Excel's answer (COM Rows(r).Height) is printed beside Oxi's
(OXI_XLSX_DUMP_ROWS); OXI_XLSX_* environment is passed through to the renderer.

    python tools/metrics/_xlsx_charset_row_height.py
"""
import os, re, subprocess, sys
from pathlib import Path
import openpyxl
from openpyxl.styles import Font

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "xlsx_probes" / "charset_rows"
FACES = [("Arial Black", 11, False), ("Arial Narrow", 8, True), ("Arial", 10, False), ("Arial", 10, True),
         ("Helvetica", 10, True), ("Lucida Sans Unicode", 12, True), ("Calibri", 10, True), ("Calibri", 11, False),
         ("Aptos Display", 10, False), ("Times New Roman", 12, False), ("Verdana", 9, False), ("Tahoma", 14, True)]
CHARSETS = [None, 0, 128, 134]


def main():
    import win32com.client
    OUT.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active
    arms = []
    r = 1
    for face, size, bold in FACES:
        for cs in CHARSETS:
            ws.cell(r, 1, "Ag").font = Font(name=face, size=size, bold=bold, charset=cs)
            arms.append((r, face, size, bold, cs)); r += 1
    path = OUT / "charset_rows.xlsx"; wb.save(path)
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        b = xl.Workbooks.Open(str(path), 0, True); sh = b.Worksheets(1)
        excel = {a[0]: round(sh.Rows(a[0]).Height / 0.75) for a in arms}
        b.Close(False)
    finally:
        xl.Quit()
    renderer = Path(os.environ.get("OXI_XLSX_RENDERER", REPO / "tools/oxi-xlsx-renderer/target/release/oxi-xlsx-renderer.exe"))
    run = subprocess.run([str(renderer), str(path), os.environ["TEMP"] + r"\_csr.png"], capture_output=True, text=True,
                         env=dict(os.environ, OXI_XLSX_DUMP_ROWS="1", OXI_XLSX_RANGE=f"1,1,{r-1},1"))
    oxi = {int(m.group(1)): int(float(m.group(2))) for m in re.finditer(r"row (\d+) px ([\d.]+)", run.stdout)}
    bad = 0
    for row, face, size, bold, cs in arms:
        o = oxi.get(row); flag = "" if o == excel[row] else "  <-"
        bad += bool(flag)
        print(f"{face:20s}{size:3}{' b' if bold else '  '} cs={str(cs):4s} Excel {excel[row]:3d}  Oxi {o}{flag}")
    print(f"{len(arms)-bad}/{len(arms)} agree")


if __name__ == "__main__":
    main()
