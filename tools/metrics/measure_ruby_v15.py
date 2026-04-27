"""V15 — extreme usWinAscent ratio verification.

Round 11 V14 covered usWinAsc/upem ratios 0.8594–0.9951.
V15 extends to 1.0601 (Meiryo) and 1.0791 (Yu Gothic UI) to verify
the corrected formula at extreme ascent ratios where the predicted
expansion drops to 2.1–2.4pt vs MS Mincho 5.2pt at base=14pt.

Writes pipeline_data/ruby_v15_measurements.json.
"""
import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = os.path.abspath("pipeline_data/docx")
OUT_PATH = os.path.abspath("pipeline_data/ruby_v15_measurements.json")

# (file_suffix, font_family_label, predicted_usWinAsc_ratio)
V15_FONTS = [
    ("YuGothicUI", "YuGothicUI_extreme", 1.0791),
    ("Meiryo_jp",  "Meiryo_extreme",     1.0601),
    ("Meiryo_en",  "Meiryo_extreme",     1.0601),
    ("MeiryoUI",   "Meiryo_extreme",     1.0601),
]

CELLS = [
    (2, 3, "default",                None, None),
    (4, 5, "raise=12halfpt(6pt)",    6.0,  None),
    (6, 7, "raise=24halfpt(12pt)",   12.0, None),
    (8, 9, "hps=base(14pt)",         None, "base"),
]

BASE_PT = 14.0
DEFAULT_HPS_PT = 7.0


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
        paras.append({"i": pi, "y_pt": y, "text": text[:80]})
    doc.Close(SaveChanges=False)
    return paras


def main() -> None:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out: dict = {
        "_meta": {
            "tool": "tools/metrics/measure_ruby_v15.py",
            "purpose": "Verify round 11 usWinAscent formula at extreme ratios (1.06+)",
            "base_pt": BASE_PT,
            "default_hps_pt": DEFAULT_HPS_PT,
        },
        "fixtures": {},
    }
    try:
        for suffix, family_label, ratio in V15_FONTS:
            fname = f"RUBY_V15_{suffix}_140dpt"
            print(f"\n=== {fname} (family={family_label}, usWin ratio={ratio}) ===")
            docx_path = os.path.join(DOCX_DIR, fname + ".docx")
            paras = measure_doc(word, docx_path)
            no_ruby_obs = []
            for (a, b) in [(1, 2), (3, 4), (5, 6), (7, 8)]:
                if a <= len(paras) and b <= len(paras):
                    pa, pb = paras[a-1], paras[b-1]
                    if pa["y_pt"] is not None and pb["y_pt"] is not None:
                        no_ruby_obs.append(round(pb["y_pt"] - pa["y_pt"], 3))
            no_ruby_lh_used = round(sum(no_ruby_obs) / len(no_ruby_obs), 3) if no_ruby_obs else BASE_PT * 9.0/7.0

            cell_results = []
            for ruby_idx, next_idx, label, raise_pt, hps_marker in CELLS:
                if next_idx > len(paras):
                    print(f"  cell {label}: SKIP")
                    continue
                rp, np_ = paras[ruby_idx-1], paras[next_idx-1]
                if rp["y_pt"] is None or np_["y_pt"] is None:
                    continue
                dy = np_["y_pt"] - rp["y_pt"]
                exp = dy - no_ruby_lh_used
                hps_pt_used = BASE_PT if hps_marker == "base" else DEFAULT_HPS_PT
                pred = max(0.0, (raise_pt or 0.0) + 0.75 * hps_pt_used - BASE_PT * ratio) if raise_pt is not None else None
                cell_results.append({
                    "label": label,
                    "dy_pt": round(dy, 3),
                    "no_ruby_lh_used": no_ruby_lh_used,
                    "expansion_pt": round(exp, 3),
                    "explicit_raise_pt": raise_pt,
                    "hps_pt": hps_pt_used,
                    "uswin_predicted_exp_pt": round(pred, 3) if pred is not None else None,
                    "delta_vs_pred_pt": round(exp - pred, 3) if pred is not None else None,
                })
                if pred is not None:
                    flag = "OK" if abs(exp - pred) <= 0.55 else "** OUTLIER **"
                    print(f"  cell {label}: dy={dy:.3f} exp={exp:.3f} usWin_pred={pred:.3f} Δ={exp - pred:+.3f}  {flag}")
                else:
                    print(f"  cell {label}: dy={dy:.3f} exp={exp:.3f} (default raise)")

            out["fixtures"][fname] = {
                "family_label": family_label,
                "uswin_ratio": ratio,
                "no_ruby_lh_observed_pt": no_ruby_obs,
                "no_ruby_lh_used_pt": no_ruby_lh_used,
                "all_paragraphs": [{"i": p["i"], "y_pt": round(p["y_pt"], 3) if p["y_pt"] else None} for p in paras],
                "cells": cell_results,
            }
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT_PATH}")


if __name__ == "__main__":
    main()
