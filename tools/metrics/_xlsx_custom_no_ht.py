# -*- coding: utf-8 -*-
"""Probe: a row that says customHeight="1" but states no ht.

Arms vary the sheet default (pinned or not) and the row's content font size;
Excel's own row height (COM, points) is printed next to Oxi's (--dump rows).

    python tools/metrics/_xlsx_custom_no_ht.py
"""
import os, re, subprocess, zipfile
from pathlib import Path
import openpyxl
from openpyxl.styles import Font

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "xlsx_probes" / "custnoht"
ARMS = {  # name: (sheetFormatPr attrs, font size of the content)
    "pinned17_sz20": ('defaultRowHeight="17.25" customHeight="1"', 20),
    "pinned17_sz8": ('defaultRowHeight="17.25" customHeight="1"', 8),
    "free15_sz20": ('defaultRowHeight="15"', 20),
    "free15_sz8": ('defaultRowHeight="15"', 8),
    "pinned30_sz11": ('defaultRowHeight="30" customHeight="1"', 11),
}


def build(name, fmt, size):
    wb = openpyxl.Workbook(); ws = wb.active
    ws["A1"] = "Ag"; ws["A1"].font = Font(size=size)
    ws["A2"] = "Ag"; ws["A2"].font = Font(size=size)
    tmp = OUT / f"{name}.tmp.xlsx"; wb.save(tmp)
    zin = zipfile.ZipFile(tmp); dst = OUT / f"{name}.xlsx"; zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "xl/worksheets/sheet1.xml":
            s = data.decode("utf8")
            s = re.sub(r"<sheetFormatPr[^>]*/>", f"<sheetFormatPr {fmt}/>", s)
            s = s.replace('<row r="1"', '<row r="1" customHeight="1"', 1)  # pinned, no ht
            s = re.sub(r'(<row r="\d+"[^>]*?) ht="[^"]*"', r"\1", s)
            data = s.encode("utf8")
        zout.writestr(item, data)
    zout.close(); zin.close(); tmp.unlink()
    return dst


def main():
    import win32com.client
    OUT.mkdir(parents=True, exist_ok=True)
    paths = {n: build(n, *a) for n, a in ARMS.items()}
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    renderer = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
    if not renderer.exists():
        renderer = Path(os.environ.get("OXI_XLSX_RENDERER", renderer))
    try:
        for n, p in paths.items():
            wb = xl.Workbooks.Open(str(p), 0, True); sh = wb.Worksheets(1)
            ex = [round(sh.Rows(r).Height / 0.75) for r in (1, 2, 3)]
            wb.Close(False)
            run = subprocess.run([str(renderer), str(p), os.environ["TEMP"] + r"\_cnh.png"], capture_output=True, text=True,
                                 env=dict(os.environ, OXI_XLSX_DUMP_ROWS="1", OXI_XLSX_RANGE="1,1,3,1"))
            ox = re.findall(r"px (\d+)", run.stdout)[:3]
            print(f"{n:16s} Excel px {ex}  Oxi px {ox}")
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
