"""V19 — measure no_ruby_LH at the EXACT baseline docGrid pattern.

Replicates V17/V18 layout (6 fonts × 5 base sizes) with the sectPr
pattern present in 49/51 baseline pipeline_data/docx docs:
  <w:docGrid w:linePitch="360"/>  (type attribute ABSENT, linePitch=18pt)

Compares to V17 (type=default), V18 (no docGrid), and LM0 lookup.
If V19 == V17/V18, Word's behavior is identical across all 3 scenarios
and Oxi's LM0 lookup table is incorrect for nearly the entire baseline.
"""
import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = os.path.abspath("pipeline_data/docx")
LM0_PATH = os.path.abspath("crates/oxidocs-core/src/font/data/lm0_lineauto.json")
OUT_PATH = os.path.abspath("pipeline_data/ruby_v19_baseline_docgrid.json")

V19_FIXTURES = [
    ("MSMincho_control", "MS Mincho"),
    ("YuMincho",         "Yu Mincho"),
    ("YuGothic",         "Yu Gothic R"),
    ("YuGothicUI",       "Yu Gothic UI"),
    ("Meiryo",           "Meiryo Reg"),
    ("MeiryoUI",         "Meiryo UI"),
]

BASES_PT = [9.0, 10.5, 11.0, 12.0, 14.0]

# V17/V18 measurements (V17==V18 was confirmed exactly across all cells)
V17_V18_MEAS = {
    "MS Mincho":    {9.0: 11.5, 10.5: 13.5, 11.0: 14.5, 12.0: 15.5, 14.0: 18.5},
    "Yu Mincho":    {9.0: 15.0, 10.5: 17.5, 11.0: 18.5, 12.0: 20.0, 14.0: 23.5},
    "Yu Gothic R":  {9.0: 15.0, 10.5: 17.5, 11.0: 18.5, 12.0: 20.0, 14.0: 23.5},
    "Yu Gothic UI": {9.0: 15.5, 10.5: 18.5, 11.0: 19.0, 12.0: 21.0, 14.0: 24.0},
    "Meiryo Reg":   {9.0: 17.5, 10.5: 20.5, 11.0: 21.5, 12.0: 23.5, 14.0: 27.0},
    "Meiryo UI":    {9.0: 15.0, 10.5: 17.5, 11.0: 18.0, 12.0: 20.0, 14.0: 23.0},
}


def measure_doc(word_app, docx_path: str) -> list[dict]:
    abs_path = os.path.abspath(docx_path)
    doc = word_app.Documents.Open(abs_path, ReadOnly=True)
    time.sleep(0.4)
    paras = []
    n = doc.Paragraphs.Count
    for pi in range(1, n + 1):
        p = doc.Paragraphs(pi)
        rng = p.Range
        try:
            y = rng.Information(6)
        except Exception:
            y = None
        text = (rng.Text or "").replace("\r", "").replace("\x07", "")
        paras.append({"i": pi, "y_pt": y, "text": text[:60]})
    doc.Close(SaveChanges=False)
    return paras


def main() -> None:
    lm0 = json.load(open(LM0_PATH, encoding="utf-8"))
    lm0_lookup_keys = {
        "MS Mincho": next((k for k in lm0 if "明朝" in k), None),
        "Yu Mincho": "Yu Mincho",
        "Yu Gothic R": "Yu Gothic",
        "Yu Gothic UI": "Yu Gothic",
        "Meiryo Reg": "Meiryo",
        "Meiryo UI": "Meiryo",
    }
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out: dict = {
        "_meta": {
            "tool": "tools/metrics/measure_ruby_v19.py",
            "purpose": "EXACT baseline docGrid pattern: <w:docGrid w:linePitch=360/> (type absent)",
            "compares_to": ["V17 (type=default)", "V18 (no docGrid)", "LM0 lookup"],
        },
        "fixtures": {},
    }
    try:
        for suffix, label in V19_FIXTURES:
            fname = f"RUBY_V19_{suffix}_baseline_docgrid"
            print(f"\n=== {fname} ({label}) ===")
            print(f"  base   V19_meas   V17/V18   Δ(V17)   LM0    Δ(LM0)")
            docx_path = os.path.join(DOCX_DIR, fname + ".docx")
            paras = measure_doc(word, docx_path)
            cells = []
            lm0_key = lm0_lookup_keys.get(label)
            for i, base_pt in enumerate(BASES_PT):
                p_a_idx = 2*i + 1
                p_b_idx = 2*i + 2
                if p_b_idx > len(paras):
                    continue
                p_a = paras[p_a_idx - 1]
                p_b = paras[p_b_idx - 1]
                if p_a["y_pt"] is None or p_b["y_pt"] is None:
                    continue
                dy = p_b["y_pt"] - p_a["y_pt"]
                v17 = V17_V18_MEAS.get(label, {}).get(base_pt)
                lm0_val = lm0.get(lm0_key, {}).get(f"{base_pt}") if lm0_key else None
                d17 = round(dy - v17, 3) if v17 is not None else None
                dlm0 = round(dy - lm0_val, 3) if lm0_val is not None else None
                cells.append({
                    "base_pt": base_pt,
                    "v19_dy_pt": round(dy, 3),
                    "v17_v18_meas_pt": v17,
                    "delta_vs_v17": d17,
                    "lm0_lookup_pt": lm0_val,
                    "delta_vs_lm0": dlm0,
                })
                v17_s = f"{v17:>7.2f}" if v17 is not None else "    N/A"
                d17_s = f"{d17:>+7.2f}" if d17 is not None else "    N/A"
                lm0_s = f"{lm0_val:>5.2f}" if lm0_val is not None else "  N/A"
                dlm0_s = f"{dlm0:>+7.2f}" if dlm0 is not None else "    N/A"
                print(f"  {base_pt:>5}pt  {dy:>7.2f}    {v17_s}  {d17_s}   {lm0_s}   {dlm0_s}")
            out["fixtures"][fname] = {"label": label, "cells": cells}
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT_PATH}")


if __name__ == "__main__":
    main()
