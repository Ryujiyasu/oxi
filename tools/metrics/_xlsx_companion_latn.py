# -*- coding: utf-8 -*-
"""Probe: does a Latin-face cell still build its line with the Normal font's
(the SX89 companion rule) in an English-mode Excel?

Normal is set to Calibri 11 or to 游ゴシック 11 (no scheme); one cell per row in
Arial 10 or Century 9, wrapped to 1-4 lines in a narrow column. Excel's row
heights (COM) are printed beside Oxi's (OXI_XLSX_* passed through).

    python tools/metrics/_xlsx_companion_latn.py
"""
import os, re, subprocess, zipfile
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Font

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "xlsx_probes" / "companion_latn"
TEXTS = ["ab", "ab cd", "ab cd ef", "ab cd ef gh"]


def build(normal):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.column_dimensions["A"].width = 3
    r = 1
    arms = []
    for face, size in (("Arial", 10), ("Century", 9), ("Times New Roman", 12)):
        for t in TEXTS:
            c = ws.cell(r, 1, t); c.font = Font(name=face, size=size); c.alignment = Alignment(wrap_text=True)
            arms.append((r, face, size, len(t.split()))); r += 1
    tmp = OUT / "tmp.xlsx"; wb.save(tmp)
    path = OUT / f"companion_{'yu' if normal != 'Calibri' else 'calibri'}.xlsx"
    zin = zipfile.ZipFile(tmp); zout = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "xl/styles.xml":
            s = data.decode("utf8")
            first = re.search(r"<font>.*?</font>", s, re.S).group(0)
            s = s.replace(first, f'<font><sz val="11"/><color theme="1"/><name val="{normal}"/><family val="2"/></font>', 1)
            data = s.encode("utf8")
        zout.writestr(item, data)
    zout.close(); zin.close(); tmp.unlink()
    return path, arms


def main():
    import win32com.client
    OUT.mkdir(parents=True, exist_ok=True)
    renderer = Path(os.environ.get("OXI_XLSX_RENDERER", REPO / "tools/oxi-xlsx-renderer/target/release/oxi-xlsx-renderer.exe"))
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        for normal in ("Calibri", "游ゴシック"):
            path, arms = build(normal)
            b = xl.Workbooks.Open(str(path), 0, True); sh = b.Worksheets(1)
            excel = {r: round(sh.Rows(r).Height / 0.75) for r, *_ in arms}
            b.Close(False)
            run = subprocess.run([str(renderer), str(path), os.environ["TEMP"] + r"\_cl.png"], capture_output=True, text=True,
                                 env=dict(os.environ, OXI_XLSX_DUMP_ROWS="1", OXI_XLSX_RANGE=f"1,1,{len(arms)},1"))
            oxi = {int(m.group(1)): int(float(m.group(2))) for m in re.finditer(r"row (\d+) px ([\d.]+)", run.stdout)}
            for r, face, size, n in arms:
                flag = "" if oxi.get(r) == excel[r] else "  <-"
                print(f"Normal {normal:6s} {face:16s}{size:3} words {n}  Excel {excel[r]:3d}  Oxi {oxi.get(r)}{flag}")
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
