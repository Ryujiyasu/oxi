"""Measure via Word COM the row heights of cell-spacing-collapse repros.

For each repro doc, we record:
 - height of the first table row (expected variable; that's what we test)
 - y coordinate of the first paragraph in row 2 (reference row position)

Prediction (if spacing collapse applies inside cell):
  S1 (sa=87, sb=87, 2p)    : row_h ≈ 2*(line=12 + before/after collapsed as max=4.35)
                            = 2*12 + 4*4.35 (before of p1, collapsed, after of p2)
                            Wait: 4.35+12+max(4.35,4.35)+12+4.35 = 37.4pt
                            If NOT collapse: 4.35+12+4.35+4.35+12+4.35 = 41.75pt
  S2 (sa=60, sb=120, 2p)   : collapse=max(3pt,6pt)=6pt; add=9pt; Δ=3pt
  S3 (sa=0, sb=120, 2p)    : collapse=max(0,6)=6; add=6; Δ=0 (no difference expected)
  S4 (sa=100, sb=100, 3p)  : 2 collapses, Δ=10pt
  S5 (sa=200, sb=100, 2p)  : collapse=max(10,5)=10; add=15; Δ=5pt
  S6 (1p)                  : no collapse possible (reference)
"""
import os
import json
import time
import win32com.client

REPRO_DIR = os.path.abspath("tools/metrics/cell_spacing_repro")
OUT = os.path.abspath("tools/metrics/cell_spacing_measurements.json")


def measure_one(app, docx_path):
    doc = app.Documents.Open(docx_path, ReadOnly=True)
    doc.Repaginate()
    time.sleep(0.3)
    tbl = doc.Tables(1)
    nrows = tbl.Rows.Count
    rows = []
    for i in range(1, nrows + 1):
        r = tbl.Rows(i)
        # paragraph in this row
        first_para = r.Cells(1).Range.Paragraphs(1).Range
        y_top = first_para.Information(6)  # pt from page top
        rows.append({
            "row": i,
            "first_para_y_pt": float(y_top),
        })
    # Also get the first paragraph of the row as a reference for p2
    paras = doc.Paragraphs
    all_paras = []
    for i in range(1, min(paras.Count, 20) + 1):
        rng = paras(i).Range
        try:
            y = rng.Information(6)
            page = rng.Information(3)
        except Exception:
            y = -1
            page = -1
        all_paras.append({
            "idx": i,
            "page": int(page),
            "y_pt": float(y),
            "text": rng.Text.rstrip("\r\n\x07")[:40],
        })
    doc.Close(False)
    return {
        "rows": rows,
        "first_20_paras": all_paras,
    }


def main():
    app = win32com.client.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    results = {}
    try:
        for name in sorted(os.listdir(REPRO_DIR)):
            if not name.endswith(".docx"):
                continue
            path = os.path.join(REPRO_DIR, name)
            print(f"Measuring {name}...")
            data = measure_one(app, path)
            results[name] = data
            # Print row1 info + next-row y
            r1 = data["rows"][0]
            if len(data["rows"]) >= 2:
                r2 = data["rows"][1]
                row1_height = r2["first_para_y_pt"] - r1["first_para_y_pt"]
                print(f"  row1 first-para y={r1['first_para_y_pt']:.2f}  row2 y={r2['first_para_y_pt']:.2f}  row1_height={row1_height:.2f}pt")
            # Print first 5 paragraph positions
            for p in data["first_20_paras"][:6]:
                print(f"    para[{p['idx']}] y={p['y_pt']:.2f} text={p['text']!r}")
    finally:
        app.Quit()

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
