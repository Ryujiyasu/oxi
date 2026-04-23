"""Measure collapse-condition repros. For each variant, row1 height is the signal:
  - 33pt = collapse applied
  - 37pt = no collapse
"""
import os
import json
import time
import win32com.client

REPRO_DIR = os.path.abspath("tools/metrics/collapse_cond_repro")
OUT = os.path.abspath("tools/metrics/collapse_cond_measurements.json")


def measure(app, path):
    doc = app.Documents.Open(path, ReadOnly=True)
    doc.Repaginate()
    time.sleep(0.2)
    tbl = doc.Tables(1)
    r1 = tbl.Rows(1).Cells(1).Range.Paragraphs(1).Range.Information(6)
    r2 = tbl.Rows(2).Cells(1).Range.Paragraphs(1).Range.Information(6)
    row1_h = r2 - r1
    doc.Close(False)
    return {"row1_y": r1, "row2_y": r2, "row1_h": row1_h}


def main():
    app = win32com.client.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    results = {}
    try:
        for n in sorted(os.listdir(REPRO_DIR)):
            if not n.endswith(".docx"): continue
            path = os.path.join(REPRO_DIR, n)
            d = measure(app, path)
            label = "collapse" if d["row1_h"] < 35 else "NO-collapse" if d["row1_h"] > 35 else "?"
            print(f"  {n}: row1_h={d['row1_h']:.2f}pt  [{label}]")
            results[n] = d
    finally:
        app.Quit()
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
