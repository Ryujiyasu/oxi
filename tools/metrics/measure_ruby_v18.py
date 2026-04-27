"""V18 — strict no-docGrid no_ruby_LH measurement.

Replicates V17 (6 fonts × 5 base sizes pure no-ruby paragraphs) but
with sectPr that omits the <w:docGrid> element entirely. Compares to:
  - V17 (with docGrid type=default linePitch=312)
  - LM0 lookup table (font/data/lm0_lineauto.json)

Hypothesis: if V18 == LM0, the LM0 table was measured for the strict
no-docGrid scenario. If V18 == V17, LM0 is incorrect for either scenario.

Writes pipeline_data/ruby_v18_no_docgrid.json.
"""
import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = os.path.abspath("pipeline_data/docx")
LM0_PATH = os.path.abspath("crates/oxidocs-core/src/font/data/lm0_lineauto.json")
OUT_PATH = os.path.abspath("pipeline_data/ruby_v18_no_docgrid.json")

V18_FIXTURES = [
    ("MSMincho_control", "MS Mincho",     "ＭＳ 明朝"),
    ("YuMincho",         "Yu Mincho",     "Yu Mincho"),
    ("YuGothic",         "Yu Gothic R",   "Yu Gothic"),
    ("YuGothicUI",       "Yu Gothic UI",  "Yu Gothic"),  # aliased in LM0?
    ("Meiryo",           "Meiryo Reg",    "Meiryo"),
    ("MeiryoUI",         "Meiryo UI",     "Meiryo"),     # aliased in LM0?
]

BASES_PT = [9.0, 10.5, 11.0, 12.0, 14.0]

# V17 measured values for direct comparison
V17_MEASUREMENTS = {
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
    # Map our font label → LM0 key
    lm0_lookup_keys = {
        "MS Mincho": next((k for k in lm0 if "明朝" in k), None),
        "Yu Mincho": "Yu Mincho",
        "Yu Gothic R": "Yu Gothic",
        "Yu Gothic UI": "Yu Gothic",   # aliased
        "Meiryo Reg": "Meiryo",
        "Meiryo UI": "Meiryo",         # aliased
    }
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out: dict = {
        "_meta": {
            "tool": "tools/metrics/measure_ruby_v18.py",
            "purpose": "strict no-docGrid no_ruby_LH; compare to V17 (with docGrid) and LM0 lookup",
            "bases_pt": BASES_PT,
        },
        "fixtures": {},
    }
    try:
        for suffix, label, _ in V18_FIXTURES:
            fname = f"RUBY_V18_{suffix}_no_docgrid"
            print(f"\n=== {fname} ({label}) ===")
            docx_path = os.path.join(DOCX_DIR, fname + ".docx")
            paras = measure_doc(word, docx_path)
            cells = []
            lm0_key = lm0_lookup_keys.get(label)
            print(f"  base   V18_meas   V17_meas    Δ(V17)   LM0    Δ(LM0)")
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
                v17 = V17_MEASUREMENTS.get(label, {}).get(base_pt)
                lm0_val = lm0.get(lm0_key, {}).get(f"{base_pt}") if lm0_key else None
                delta_v17 = round(dy - v17, 3) if v17 is not None else None
                delta_lm0 = round(dy - lm0_val, 3) if lm0_val is not None else None
                cells.append({
                    "base_pt": base_pt,
                    "v18_dy_pt": round(dy, 3),
                    "v17_meas_pt": v17,
                    "delta_vs_v17_pt": delta_v17,
                    "lm0_lookup_pt": lm0_val,
                    "delta_vs_lm0_pt": delta_lm0,
                })
                v17_s = f"{v17:>7.2f}" if v17 is not None else "    N/A"
                d17_s = f"{delta_v17:>+7.2f}" if delta_v17 is not None else "    N/A"
                lm0_s = f"{lm0_val:>5.2f}" if lm0_val is not None else "  N/A"
                dlm0_s = f"{delta_lm0:>+7.2f}" if delta_lm0 is not None else "    N/A"
                print(f"  {base_pt:>5}pt  {dy:>7.2f}    {v17_s}    {d17_s}   {lm0_s}   {dlm0_s}")
            out["fixtures"][fname] = {
                "label": label,
                "lm0_key": lm0_key,
                "cells": cells,
            }
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT_PATH}")


if __name__ == "__main__":
    main()
