# -*- coding: utf-8 -*-
"""Collect Excel's own row height for one cell of font F at size S, across many
installed faces and sizes, with the Normal font shrunk to 1pt so the sheet
default never props a row open. Writes rows of (face, size, bold, excel_px)
to pipeline_data/xlsx_probes/row_metrics/excel_rows.json for fitting against
GDI text metrics.

    python tools/metrics/_xlsx_row_height_by_metrics.py
"""
import json, re, zipfile
from pathlib import Path
import openpyxl
from openpyxl.styles import Font

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "xlsx_probes" / "row_metrics"
FACES = ["Arial", "Arial Black", "Arial Narrow", "Calibri", "Calibri Light", "Cambria", "Candara", "Consolas",
         "Constantia", "Corbel", "Courier New", "Franklin Gothic Medium", "Gabriola", "Georgia", "Impact",
         "Lucida Console", "Lucida Sans Unicode", "Palatino Linotype", "Segoe UI", "Segoe UI Light", "Tahoma",
         "Times New Roman", "Trebuchet MS", "Verdana", "Century Gothic", "Garamond", "Book Antiqua", "Bahnschrift",
         "Aptos", "Aptos Narrow", "Aptos Display", "Comic Sans MS", "Ebrima", "Leelawadee UI", "Microsoft Sans Serif",
         "MS Reference Sans Serif", "Sitka Text", "Sylfaen", "Ink Free", "Dubai"]
SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36]
BOLD = (False, True)
NAME = "excel_rows.json"


def main():
    import sys
    import win32com.client
    global FACES, SIZES, BOLD, NAME
    if "--sweep" in sys.argv:
        # Every whole pixel of em from 8 to 96 (0.75pt apart), regular only,
        # for the faces the measured table does not hold.
        FACES = [f for f in sys.argv[sys.argv.index("--sweep") + 1].split(",")]
        SIZES = [px * 0.75 for px in range(8, 97)]
        BOLD = (False,)
        NAME = "excel_rows_sweep.json"
    OUT.mkdir(parents=True, exist_ok=True)
    arms = [(face, size, bold) for face in FACES for size in SIZES for bold in BOLD]
    # A workbook holds about a thousand distinct fonts; past that Excel stops
    # deriving the rows and every later one reads 25px. Measure in chunks.
    got = []
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        for start in range(0, len(arms), 600):
            chunk = arms[start:start + 600]
            wb = openpyxl.Workbook(); ws = wb.active
            for r, (face, size, bold) in enumerate(chunk, 1):
                ws.cell(r, 1, "Ag").font = Font(name=face, size=size, bold=bold)
            tmp = OUT / "rows.tmp.xlsx"; wb.save(tmp)
            path = OUT / f"rows_{start}.xlsx"
            zin = zipfile.ZipFile(tmp); zout = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/styles.xml":
                    t = data.decode("utf8")
                    t = re.sub(r'(<fonts[^>]*>\s*<font>.*?<sz val=")[^"]*', r"\g<1>1", t, count=1, flags=re.S)
                    data = t.encode("utf8")
                zout.writestr(item, data)
            zout.close(); zin.close(); tmp.unlink()
            b = xl.Workbooks.Open(str(path), 0, True); sh = b.Worksheets(1)
            got += [(face, size, bold, round(sh.Rows(r).Height / 0.75)) for r, (face, size, bold) in enumerate(chunk, 1)]
            empty = round(sh.Rows(len(chunk) + 5).Height / 0.75)
            b.Close(False)
    finally:
        xl.Quit()
    (OUT / NAME).write_text(json.dumps({"empty_row_px": empty, "rows": got}, ensure_ascii=False))
    print(f"{len(got)} rows, empty row {empty}px")


if __name__ == "__main__":
    main()
