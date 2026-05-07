"""COM-measure Word's actual line/row heights for grid-snap minimal repros.

For each L1-L8 doc, measure each paragraph's Y position, compute inter-
paragraph ΔY (= effective line height), and dump to JSON. This establishes
Word's ground-truth behavior for the line-height category.

Output: pipeline_data/grid_snap_repros_word.json
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
OUT = os.path.join(REPO, "pipeline_data", "grid_snap_repros_word.json")


def measure_doc(word, docx_path: str) -> dict:
    label = os.path.splitext(os.path.basename(docx_path))[0]
    print(f"=== {label} ===")
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    out = {"label": label, "paragraphs": []}
    try:
        sec = doc.Sections(1)
        out["page_margin_top_pt"] = round(sec.PageSetup.TopMargin, 3)
        out["doc_grid_line_pitch_tw"] = sec.PageSetup.LinePitch if hasattr(sec.PageSetup, "LinePitch") else None

        n = doc.Paragraphs.Count
        for i in range(1, n + 1):
            try:
                para = doc.Paragraphs(i)
                rng = para.Range
                first = doc.Range(rng.Start, rng.Start)
                y = first.Information(WD_VPOS)
                x = first.Information(WD_HPOS)
                pg = first.Information(WD_PAGE)
                txt = rng.Text
                fmt = para.Format
                line_spacing = fmt.LineSpacing
                line_spacing_rule = fmt.LineSpacingRule
                out["paragraphs"].append({
                    "i": i,
                    "page": pg,
                    "x": round(x, 3),
                    "y": round(y, 3),
                    "text": txt[:30],
                    "line_spacing": round(line_spacing, 3) if line_spacing is not None else None,
                    "line_spacing_rule": line_spacing_rule,
                })
            except Exception as e:
                out["paragraphs"].append({"i": i, "error": str(e)})

        # Compute deltas
        prev_y = None
        prev_pg = None
        deltas = []
        for p in out["paragraphs"]:
            if "y" not in p:
                continue
            if prev_y is not None and p["page"] == prev_pg:
                d = p["y"] - prev_y
                deltas.append({"i": p["i"], "delta_y": round(d, 3)})
            prev_y = p["y"]
            prev_pg = p["page"]
        out["deltas"] = deltas
        # Stats
        if deltas:
            ds = [d["delta_y"] for d in deltas]
            out["delta_mean"] = round(sum(ds) / len(ds), 3)
            out["delta_min"] = min(ds)
            out["delta_max"] = max(ds)

        # Display
        for p in out["paragraphs"][:10]:
            if "y" in p:
                print(f"  i={p['i']} pg={p['page']} y={p['y']:.2f} txt={p['text']!r}")
        print(f"  Inter-para ΔY: mean={out.get('delta_mean')} min={out.get('delta_min')} max={out.get('delta_max')}")
    finally:
        doc.Close(SaveChanges=False)
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    docx_paths = sorted(glob.glob(os.path.join(REPRO_DIR, "*.docx")))
    print(f"Measuring {len(docx_paths)} repros...")

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    results = []
    try:
        for dp in docx_paths:
            try:
                results.append(measure_doc(word, dp))
            except Exception as e:
                print(f"  ERROR {os.path.basename(dp)}: {e}")
                results.append({"label": os.path.basename(dp), "error": str(e)})
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"results": results}, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")

    # Summary
    print("\n=== Summary ===")
    print(f"{'label':<6} {'pitch_tw':<10} {'mean_dy':<10} {'min':<8} {'max':<8}")
    for r in results:
        if "delta_mean" in r:
            print(f"{r['label']:<6} {r.get('doc_grid_line_pitch_tw','?'):<10} {r['delta_mean']:<10} {r['delta_min']:<8} {r['delta_max']:<8}")


if __name__ == "__main__":
    main()
