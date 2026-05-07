"""COM-measure Word's actual cell line height for b5f706_V{1..6} variants.

For each variant, measure each cell paragraph's Y position and compute
inter-paragraph delta_y = effective line height. Compare across variants
to identify which factor (font, font_size, linePitch, orientation, snap)
flips Word's cell-internal grid-snap behavior.

Output: pipeline_data/b5f706_l7_variants_word.json
"""
from __future__ import annotations

import glob
import json
import os
import sys

import win32com.client

WD_HPOS = 5
WD_VPOS = 6
WD_PAGE = 3

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
REPRO_DIR = os.path.join(REPO, "tools", "golden-test", "repros", "grid_snap")
OUT = os.path.join(REPO, "pipeline_data", "b5f706_l7_variants_word.json")


def measure_doc(word, docx_path: str) -> dict:
    label = os.path.splitext(os.path.basename(docx_path))[0]
    print(f"=== {label} ===")
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    out = {"label": label, "paragraphs": []}
    try:
        sec = doc.Sections(1)
        out["page_margin_top_pt"] = round(sec.PageSetup.TopMargin, 3)
        try:
            out["doc_grid_line_pitch_tw"] = sec.PageSetup.LinePitch
        except Exception:
            out["doc_grid_line_pitch_tw"] = None

        n = doc.Paragraphs.Count
        out["paragraph_count"] = n
        for i in range(1, n + 1):
            try:
                para = doc.Paragraphs(i)
                rng = para.Range
                first = doc.Range(rng.Start, rng.Start)
                y = first.Information(WD_VPOS)
                x = first.Information(WD_HPOS)
                pg = first.Information(WD_PAGE)
                txt = rng.Text
                out["paragraphs"].append({
                    "i": i,
                    "page": pg,
                    "x": round(x, 3),
                    "y": round(y, 3),
                    "text": txt[:30],
                })
            except Exception as e:
                out["paragraphs"].append({"i": i, "error": str(e)})

        # Compute delta_y between consecutive same-page paragraphs
        deltas = []
        for i in range(1, len(out["paragraphs"])):
            cur = out["paragraphs"][i]
            prev = out["paragraphs"][i-1]
            if "y" in cur and "y" in prev and cur.get("page") == prev.get("page"):
                deltas.append({
                    "i": cur["i"],
                    "delta_y": round(cur["y"] - prev["y"], 3),
                    "text": cur["text"][:20],
                })
        out["deltas"] = deltas
        if deltas:
            ds = [d["delta_y"] for d in deltas if d["delta_y"] > 0]
            if ds:
                out["delta_mean_positive"] = round(sum(ds) / len(ds), 3)
                out["delta_min_positive"] = min(ds)
                out["delta_max_positive"] = max(ds)
    finally:
        doc.Close(SaveChanges=False)
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    docx_files = sorted(glob.glob(os.path.join(REPRO_DIR, "b5f706_V*.docx")))
    if not docx_files:
        print(f"No b5f706_V*.docx found in {REPRO_DIR}")
        sys.exit(1)
    print(f"Found {len(docx_files)} docx files")

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for path in docx_files:
            r = measure_doc(word, path)
            results.append(r)
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"results": results}, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")

    # Print compact summary
    print("\n=== Summary: cell line height (paragraph-to-paragraph deltaY) ===")
    print(f"{'label':<28s} {'pitch_tw':>9s} {'mean_dy':>9s} {'min':>6s} {'max':>6s}")
    for r in results:
        print(f"{r['label']:<28s} {r.get('doc_grid_line_pitch_tw', '-'):>9} "
              f"{r.get('delta_mean_positive', '-'):>9} "
              f"{r.get('delta_min_positive', '-'):>6} "
              f"{r.get('delta_max_positive', '-'):>6}")


if __name__ == "__main__":
    main()
