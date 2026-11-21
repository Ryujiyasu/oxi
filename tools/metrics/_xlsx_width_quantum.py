# -*- coding: utf-8 -*-
"""Probe: how Excel turns a stated column width with a fraction of 1/256
into pixels -- truncated to 1/256 first, rounded, or used as is.

Widths are chosen so 256*w has fractions .1 .3 .5 .7 .9 at several bases,
near the px boundaries of Calibri 11 (digit 7, padding 18/256).

    python tools/metrics/_xlsx_width_quantum.py
"""
import os, re, subprocess
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "xlsx_probes" / "width_quantum"


def main():
    import win32com.client
    OUT.mkdir(parents=True, exist_ok=True)
    widths = []
    for units in (2615, 2633, 7991, 1500, 3000):
        for frac in (0.1, 0.3, 0.5, 0.7, 0.9, 0.99):
            widths.append((units + frac) / 256)
    wb = openpyxl.Workbook(); ws = wb.active
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        ws.cell(1, i, "x")
    path = OUT / "width_quantum.xlsx"; wb.save(path)
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        b = xl.Workbooks.Open(str(path), 0, True); sh = b.Worksheets(1)
        excel = [round(sh.Columns(i).Width / 0.75) for i in range(1, len(widths) + 1)]
        b.Close(False)
    finally:
        xl.Quit()
    for w, px in zip(widths, excel):
        u = 256 * w
        f = lambda q: int(((q + 18) / 256) * 7)
        print(f"256w={u:9.2f}  Excel {px}  trunc {f(int(u))}  round {f(round(u))}  raw {f(u)}")


if __name__ == "__main__":
    main()
