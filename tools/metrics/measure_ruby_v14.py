"""V14 (font-family × ruby) focused COM measurement.

Tests round 10 TTF prediction:
  ascent_constant = base × OS/2.sTypoAscender / unitsPerEm

For 5 fonts at base=14pt with V13-pattern 9-paragraph layout, measure
ruby expansion per cell and compare to TTF-predicted values.

Predicted expansions at base=14pt:
  cell             MS_legacy(0.8594)  Yu/BIZ_std(0.8799)  Δ
  default          ?                   ?                   varies (default raise opacity)
  raise=12halfpt   max(0, 6+5.25-12.031)=0    max(0, 6+5.25-12.319)=0   0
  raise=24halfpt   max(0, 12+5.25-12.031)=5.22  max(0, 12+5.25-12.319)=4.93  +0.29 (signal cell!)
  hps=base         high (default raise dependent)

Writes pipeline_data/ruby_v14_measurements.json.
"""
import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = os.path.abspath("pipeline_data/docx")
OUT_PATH = os.path.abspath("pipeline_data/ruby_v14_measurements.json")

# (file_suffix, font_family_label, predicted_ratio)
V14_FONTS = [
    ("MSMincho_control", "MS_legacy", 0.8594),
    ("YuMincho",         "Yu_BIZ_std", 0.8799),
    ("YuGothic_jp",      "Yu_BIZ_std", 0.8799),
    ("YuGothic_en",      "Yu_BIZ_std", 0.8799),
    ("BIZUDMincho",      "Yu_BIZ_std", 0.8799),
]

CELLS = [
    # (ruby_para_idx, next_para_idx, label, explicit_raise_pt_or_None, hps_marker)
    (2, 3, "default",                None, None),
    (4, 5, "raise=12halfpt(6pt)",    6.0,  None),
    (6, 7, "raise=24halfpt(12pt)",   12.0, None),  # signal cell
    (8, 9, "hps=base(14pt)",         None, "base"),
]

BASE_PT = 14.0
DEFAULT_HPS_PT = 7.0
NO_RUBY_LH_PRED = BASE_PT * 9.0 / 7.0  # = 18.0pt


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
            "tool": "tools/metrics/measure_ruby_v14.py",
            "purpose": "Verify round 10 TTF-derived ascent prediction across font families",
            "base_pt": BASE_PT,
            "default_hps_pt": DEFAULT_HPS_PT,
            "no_ruby_lh_predicted_pt": NO_RUBY_LH_PRED,
        },
        "fixtures": {},
    }
    try:
        for suffix, family_label, ratio in V14_FONTS:
            fname = f"RUBY_V14_{suffix}_140dpt"
            print(f"\n=== {fname} (family={family_label}, ratio={ratio}) ===")
            docx_path = os.path.join(DOCX_DIR, fname + ".docx")
            paras = measure_doc(word, docx_path)
            no_ruby_obs = []
            for (a, b) in [(1, 2), (3, 4), (5, 6), (7, 8)]:
                if a <= len(paras) and b <= len(paras):
                    pa, pb = paras[a-1], paras[b-1]
                    if pa["y_pt"] is not None and pb["y_pt"] is not None:
                        no_ruby_obs.append(round(pb["y_pt"] - pa["y_pt"], 3))
            no_ruby_lh_used = round(sum(no_ruby_obs) / len(no_ruby_obs), 3) if no_ruby_obs else NO_RUBY_LH_PRED

            cell_results = []
            for ruby_idx, next_idx, label, raise_pt, hps_marker in CELLS:
                if next_idx > len(paras):
                    print(f"  cell {label}: SKIP (idx out of range)")
                    continue
                rp, np_ = paras[ruby_idx-1], paras[next_idx-1]
                if rp["y_pt"] is None or np_["y_pt"] is None:
                    print(f"  cell {label}: SKIP (None y)")
                    continue
                dy = np_["y_pt"] - rp["y_pt"]
                exp = dy - no_ruby_lh_used
                hps_pt_used = BASE_PT if hps_marker == "base" else DEFAULT_HPS_PT
                ttf_pred = max(0.0, (raise_pt or 0.0) + 0.75 * hps_pt_used - BASE_PT * ratio) if raise_pt is not None else None
                cell_results.append({
                    "label": label,
                    "ruby_para_idx": ruby_idx,
                    "next_para_idx": next_idx,
                    "dy_pt": round(dy, 3),
                    "no_ruby_lh_used": no_ruby_lh_used,
                    "expansion_pt": round(exp, 3),
                    "explicit_raise_pt": raise_pt,
                    "hps_pt": hps_pt_used,
                    "ttf_predicted_exp_pt": round(ttf_pred, 3) if ttf_pred is not None else None,
                    "delta_vs_ttf_pt": round(exp - ttf_pred, 3) if ttf_pred is not None else None,
                })
                if ttf_pred is not None:
                    print(
                        f"  cell {label}: dy={dy:.3f} exp={exp:.3f} "
                        f"ttf_pred={ttf_pred:.3f} Δ={exp - ttf_pred:+.3f}"
                    )
                else:
                    print(f"  cell {label}: dy={dy:.3f} exp={exp:.3f} (default raise — no closed-form prediction)")

            out["fixtures"][fname] = {
                "family_label": family_label,
                "ttf_ratio": ratio,
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
