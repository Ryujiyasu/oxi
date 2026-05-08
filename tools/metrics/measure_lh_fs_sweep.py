"""Measure Word actual paragraph dy (= lh) for lh_M_fs*_LM*_*.docx via COM.

Output JSON with per-doc delta_y mean and Oxi comparison.
"""
from __future__ import annotations

import glob
import json
import os
import sys

import win32com.client

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
REPRO_DIR = os.path.join(REPO, "tools", "golden-test", "repros", "grid_snap")
OUT = os.path.join(REPO, "pipeline_data", "lh_fs_sweep_word.json")

WD_VPOS = 6


def measure(word, docx_path: str) -> dict:
    label = os.path.splitext(os.path.basename(docx_path))[0]
    doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False)
    out = {"label": label, "paragraphs": []}
    try:
        n = doc.Paragraphs.Count
        for i in range(1, n + 1):
            try:
                p = doc.Paragraphs(i)
                rng = p.Range
                first = doc.Range(rng.Start, rng.Start)
                y = first.Information(WD_VPOS)
                txt = rng.Text[:30]
                out["paragraphs"].append({"i": i, "y": round(y, 3), "text": txt})
            except Exception as e:
                out["paragraphs"].append({"i": i, "error": str(e)})
        # delta_y
        ys = [p["y"] for p in out["paragraphs"] if "y" in p]
        if len(ys) >= 2:
            dys = [round(ys[i] - ys[i-1], 3) for i in range(1, min(7, len(ys)))]
            positive_dys = [d for d in dys if 0 < d < 50]
            if positive_dys:
                out["delta_y_first_few"] = dys
                out["delta_y_mean"] = round(sum(positive_dys) / len(positive_dys), 3)
    finally:
        doc.Close(SaveChanges=False)
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    docx_files = sorted(glob.glob(os.path.join(REPRO_DIR, "lh_M_fs*_*.docx")))
    print(f"Found {len(docx_files)} files")

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        for path in docx_files:
            try:
                r = measure(word, path)
                results.append(r)
                print(f"  {r['label']}: dy_mean = {r.get('delta_y_mean','-')}")
            except Exception as e:
                print(f"  ERROR on {path}: {e}")
                results.append({"label": os.path.basename(path), "error": str(e)})
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"results": results}, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT}")

    # Compact summary table
    print(f"\n{'label':<28s} {'paras':>5s} {'mean_dy':>8s} {'first_3_dy':>20s}")
    for r in sorted(results, key=lambda x: x['label']):
        if 'delta_y_mean' in r:
            dys = r.get('delta_y_first_few', [])
            print(f"{r['label']:<28s} {len(r['paragraphs']):>5d} {r['delta_y_mean']:>8.3f} {str(dys[:3]):>20s}")


if __name__ == "__main__":
    main()
