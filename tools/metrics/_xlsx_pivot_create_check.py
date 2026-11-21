# -*- coding: utf-8 -*-
"""Pivot creation: for every truth arm (_xlsx_pivot_gen.py), Oxi creates the
same pivot on a data-only workbook and saves it; Excel then opens the file
(failing or repairing counts as a failure), and its pivot -- refreshed on
load -- and the cells Oxi wrote are both held against the truth grid.

    python tools/metrics/_xlsx_pivot_create_check.py
"""
import json, os, subprocess, sys, tempfile
from pathlib import Path
import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _xlsx_pivot_gen as gen  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
RENDERER = Path(os.environ.get("OXI_XLSX_RENDERER", REPO / "tools/oxi-xlsx-renderer/target/release/oxi-xlsx-renderer.exe"))
FIELDS = ["Region", "Product", "Month", "Sales", "Qty"]
HOW = {gen.XL_SUM: "sum", gen.XL_COUNT: "count", gen.XL_AVG: "average", gen.XL_MAX: "max", gen.XL_MIN: "min"}


def retry(fn, tries=50):
    import time
    for k in range(tries):
        try:
            return fn()
        except Exception as bad:
            if "-2147418111" not in str(bad) or k == tries - 1:
                raise
            time.sleep(0.2)


def norm(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, float):
        return round(v, 9)
    return v


def main():
    work = Path(tempfile.mkdtemp(prefix="pivot_create_"))
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    rows = gen.data_rows()
    ok = total = 0
    try:
        groups = [(gen.ARMS, REPO / "pipeline_data/xlsx_pivot"), (gen.ARMS2, REPO / "pipeline_data/xlsx_pivot")]
        for arms, truth_dir in groups:
            for name, arm in arms.items():
                if arm.get("hide") or arm.get("page_pick") or arm.get("data_on_rows"):
                    continue  # creation takes fields and functions only
                # data-only book, with the same blanks the truth had
                wb = xl.Workbooks.Add(); data = wb.Worksheets(1); data.Name = "Data"
                for j, h in enumerate(FIELDS, 1):
                    data.Cells(1, j).Value = h
                for i, r in enumerate(rows, 2):
                    for j, v in enumerate(r, 1):
                        if arm.get("blanks") and (i * 7 + j) % 11 == 0:
                            continue
                        data.Cells(i, j).Value = v
                pv = wb.Worksheets.Add(After=data); pv.Name = "Pivot"
                base = work / f"{name}.base.xlsx"; wb.SaveAs(str(base), 51); wb.Close(False)
                spec = {"name": "PT1", "source_sheet": "Data", "source_from": [0, 0], "source_to": [len(rows), 4],
                        "target_sheet": "Pivot", "at": [2, 0],
                        "rows": [FIELDS.index(f) for f in arm.get("rows", [])],
                        "cols": [FIELDS.index(f) for f in arm.get("cols", [])],
                        "pages": [FIELDS.index(f) for f in arm.get("pages", [])],
                        "data": [[FIELDS.index(f), HOW[fn]] for f, fn in arm.get("data", [])]}
                spec_path = work / f"{name}.json"; spec_path.write_text(json.dumps(spec))
                out = work / f"{name}.xlsx"
                run = subprocess.run([str(RENDERER), str(base), "--add-pivot", str(out), str(spec_path)], capture_output=True, text=True)
                total += 1
                if not out.exists():
                    print(f"  {name:26s} SAVE FAILED {run.stderr[:150]}")
                    continue
                truth = json.loads((truth_dir / f"{name}.truth.json").read_text(encoding="cp932"))
                want = [[norm(v) for v in r] for r in truth["grid"]]
                try:
                    wb = xl.Workbooks.Open(str(out), 0, True, None, None, None, True, None, None, None, None, None, None, None, 0)
                except Exception as bad:
                    print(f"  {name:26s} EXCEL WOULD NOT OPEN {str(bad)[:100]}")
                    continue
                sh = wb.Worksheets("Pivot")
                n_pt = retry(lambda: sh.PivotTables().Count)
                h, w = len(want), max(len(r) for r in want)
                got = retry(lambda: [[norm(sh.Cells(truth["first_row"] + i, truth["first_col"] + j).Value) for j in range(w)] for i in range(h)])
                if n_pt:
                    retry(lambda: sh.PivotTables(1).RefreshTable())
                    excel = retry(lambda: [[norm(sh.Cells(truth["first_row"] + i, truth["first_col"] + j).Value) for j in range(w)] for i in range(h)])
                else:
                    excel = None
                wb.Close(False)
                want_p = [r + [None] * (w - len(r)) for r in want]
                good = n_pt == 1 and got == want_p and excel == want_p
                ok += good
                print(f"  {name:26s} {'OK' if good else 'DIFF'}  pivots={n_pt} oxi_cells={'ok' if got == want_p else 'diff'} excel_refresh={'ok' if excel == want_p else 'diff'}")
                if excel and excel != want_p:
                    for a, b in zip(want_p, excel):
                        if a != b:
                            print(f"      truth {a}"); print(f"      excel {b}")
    finally:
        xl.Quit()
    print(f"{ok}/{total} created pivots open in Excel and match")


if __name__ == "__main__":
    main()
