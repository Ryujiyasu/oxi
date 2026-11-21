# -*- coding: utf-8 -*-
"""End-to-end: edit a pivot's source in Oxi, refresh the pivots in Oxi, save;
then open the saved book in Excel, read the pivot range as saved (Oxi's
refresh), make Excel refresh the same table, and compare the two.

    python tools/metrics/_xlsx_pivot_refresh_check.py [dir]
"""
import os, subprocess, sys, tempfile
from pathlib import Path
import win32com.client

REPO = Path(__file__).resolve().parents[2]
RENDERER = Path(os.environ.get("OXI_XLSX_RENDERER", REPO / "tools/oxi-xlsx-renderer/target/release/oxi-xlsx-renderer.exe"))
# Source edits: Sales of record 1 up, a region renamed (a new item), one Qty blanked to text.
EDITS = ["Data!2,3=9999", "Data!5,0=Central", "Data!9,4=7"]


def grid(sheet, address):
    rng = sheet.Range(address)
    return [[rng.Cells(i, j).Value for j in range(1, rng.Columns.Count + 1)] for i in range(1, rng.Rows.Count + 1)]


def main():
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else REPO / "pipeline_data" / "xlsx_pivot"
    out_dir = Path(tempfile.mkdtemp(prefix="pivot_refresh_"))
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    ok = total = 0
    try:
        for book in sorted(root.glob("*.xlsx")):
            out = out_dir / book.name
            run = subprocess.run([str(RENDERER), str(book), "--refresh-pivots", str(out), *EDITS], capture_output=True, text=True)
            if not out.exists():
                print(f"  {book.stem:26s} SAVE FAILED {run.stderr[:120]}")
                continue
            total += 1
            wb = xl.Workbooks.Open(str(out), 0, False)
            sh = wb.Worksheets("Pivot"); pt = sh.PivotTables(1)
            before = pt.TableRange2.Address
            pt.RefreshTable()
            after = pt.TableRange2.Address
            span = sh.Range(before).Address if before == after else sh.Range(f"{before},{after}").Address
            # what Oxi wrote is gone after Excel's refresh, so read it first from a fresh open
            wb.Close(False)
            wb = xl.Workbooks.Open(str(out), 0, True); sh = wb.Worksheets("Pivot")
            area = sh.Range("A1").Resize(max(sh.UsedRange.Rows.Count, 1) + sh.UsedRange.Row, 20)
            mine = grid(sh, area.Address)
            wb.Close(False)
            wb = xl.Workbooks.Open(str(out), 0, True); sh = wb.Worksheets("Pivot"); sh.PivotTables(1).RefreshTable()
            area2 = sh.Range("A1").Resize(max(sh.UsedRange.Rows.Count, 1) + sh.UsedRange.Row, 20)
            theirs = grid(sh, area2.Address)
            wb.Close(False)
            h = max(len(mine), len(theirs))
            pad = lambda g: g + [[None] * 20] * (h - len(g))
            same = pad(mine) == pad(theirs)
            ok += same
            print(f"  {book.stem:26s} {'OK' if same else 'DIFF'}")
            if not same:
                for i, (a, b) in enumerate(zip(pad(theirs), pad(mine))):
                    if a != b:
                        print(f"      row {i + 1}: excel {[v for v in a if v is not None]}\n              oxi   {[v for v in b if v is not None]}")
                        break
    finally:
        xl.Quit()
    print(f"{ok}/{total} refreshed pivots match Excel's refresh")


if __name__ == "__main__":
    main()
