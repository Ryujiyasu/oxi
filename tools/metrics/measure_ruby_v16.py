"""V16 — tier B/C raise sweep at base=14pt.

For Yu Gothic UI and Meiryo Regular at base=14pt, hps=7pt fixed,
sweep raise ∈ {6, 12, 18, 24, 36}pt to characterize raise→exp slope.

Pattern (11 paragraphs per fixture):
  P0  ref
  P1  ruby raise=6pt
  P2  ref      ← dy(P1,P2) = P1's height = no_ruby_LH + exp(raise=6)
  P3  ruby raise=12pt
  P4  ref
  ...
  P9  ruby raise=36pt
  P10 ref

Linear fit: exp(r) = a × r + b. If a ≈ 1.0 → tier-A formula structure
preserved (just different intercept = different ascent constant).
If a < 1.0 → formula structure broken for tier B/C.

Writes pipeline_data/ruby_v16_measurements.json.
"""
import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = os.path.abspath("pipeline_data/docx")
OUT_PATH = os.path.abspath("pipeline_data/ruby_v16_measurements.json")

V16_FIXTURES = [
    ("YuGothicUI",        "Yu Gothic UI",     1.0791),
    ("MeiryoRegular_jp",  "Meiryo (jp)",      1.0601),
]

RAISES_PT = [6.0, 12.0, 18.0, 24.0, 36.0]
BASE_PT = 14.0
HPS_PT = 7.0


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
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out: dict = {
        "_meta": {
            "tool": "tools/metrics/measure_ruby_v16.py",
            "purpose": "Tier B/C raise sweep linearity test",
            "base_pt": BASE_PT, "hps_pt": HPS_PT,
            "raises_pt": RAISES_PT,
        },
        "fixtures": {},
    }
    try:
        for suffix, label, ratio in V16_FIXTURES:
            fname = f"RUBY_V16_{suffix}_140dpt_raisesweep"
            print(f"\n=== {fname} ({label}, usWin ratio={ratio}) ===")
            docx_path = os.path.join(DOCX_DIR, fname + ".docx")
            paras = measure_doc(word, docx_path)
            # Compute no_ruby_LH from successive ref-pair? Actually no — refs are
            # at indices 0, 2, 4, 6, 8, 10. Ruby paragraphs at 1, 3, 5, 7, 9.
            # dy(ref_i, ruby_{i+1}) = ref_i height = no_ruby_LH (consistent)
            # dy(ruby_i, ref_{i+1}) = ruby_i height = no_ruby_LH + exp_i
            # Use dy(ref → next ref) = 2 paragraph heights summed if no other.
            # Simplest: dy(ref_0, ref_2) covers ruby_1's height + ruby_1 height = no_ruby + ruby
            # Actually better: dy(ref_{i}, ref_{i+1}) = no_ruby_LH (ref) + ruby_LH (ruby_between)
            #
            # Cleanest: dy(ruby, next_ref) = ruby's own height (since ref follows immediately)
            # And dy(ref, next_ruby) = ref's own height = no_ruby_LH.
            #
            # Para indices (1-based): ref=1, ruby=2, ref=3, ruby=4, ref=5, ...
            n_paras = len(paras)
            no_ruby_dy = []
            for ref_idx in (1, 3, 5, 7, 9, 11):
                if ref_idx + 1 <= n_paras:
                    a = paras[ref_idx - 1]
                    b = paras[ref_idx]
                    if a["y_pt"] is not None and b["y_pt"] is not None:
                        no_ruby_dy.append(round(b["y_pt"] - a["y_pt"], 3))
            no_ruby_lh = sum(no_ruby_dy) / len(no_ruby_dy) if no_ruby_dy else BASE_PT * 9.0/7.0

            # Ruby paragraphs are at idx 2, 4, 6, 8, 10
            cells = []
            for k, ruby_idx in enumerate((2, 4, 6, 8, 10)):
                if ruby_idx + 1 > n_paras:
                    continue
                rp = paras[ruby_idx - 1]
                np_ = paras[ruby_idx]
                if rp["y_pt"] is None or np_["y_pt"] is None:
                    continue
                dy = np_["y_pt"] - rp["y_pt"]
                exp = dy - no_ruby_lh
                r_pt = RAISES_PT[k]
                # tier-A predicted: max(0, r_pt + 0.75*hps - base*ratio)
                pred_tierA = max(0.0, r_pt + 0.75*HPS_PT - BASE_PT * ratio)
                cells.append({
                    "raise_pt": r_pt,
                    "dy_pt": round(dy, 3),
                    "exp_pt": round(exp, 3),
                    "tierA_pred_pt": round(pred_tierA, 3),
                    "delta_vs_tierA_pt": round(exp - pred_tierA, 3),
                })

            # Linear fit: exp = a*raise + b
            if len(cells) >= 2:
                xs = [c["raise_pt"] for c in cells]
                ys = [c["exp_pt"] for c in cells]
                n = len(xs)
                mean_x = sum(xs)/n
                mean_y = sum(ys)/n
                num = sum((xs[i]-mean_x)*(ys[i]-mean_y) for i in range(n))
                den = sum((xs[i]-mean_x)**2 for i in range(n))
                slope = num/den if den else 0
                intercept = mean_y - slope*mean_x
                # constant C such that exp = max(0, slope*raise + 0.75*hps - C)
                # at typical model: exp = raise + 0.75*hps - asc → slope=1, intercept=0.75*hps - asc
                # If slope ≠ 1, formula structure differs.
                C_implied = 0.75*HPS_PT - intercept
                print(f"  no_ruby_LH = {no_ruby_lh:.3f} pt (mean of {len(no_ruby_dy)} ref-pair dy)")
                print(f"  raise(pt)→exp(pt): " + ", ".join(f"{c['raise_pt']:g}→{c['exp_pt']:.3f}" for c in cells))
                print(f"  linear fit: exp = {slope:.4f} × raise + {intercept:.3f}")
                print(f"  if formula = max(0, slope*raise + 0.75*hps - C), C = {C_implied:.3f}")
                print(f"  TIER-A formula assumes slope = 1.0; observed = {slope:.4f}")
                # Per-cell tier-A delta:
                for c in cells:
                    print(f"    r={c['raise_pt']:g}: tierA_pred={c['tierA_pred_pt']:.3f}, meas={c['exp_pt']:.3f}, Δ={c['delta_vs_tierA_pt']:+.3f}")

                out["fixtures"][fname] = {
                    "label": label,
                    "uswin_ratio": ratio,
                    "no_ruby_lh_pt": round(no_ruby_lh, 3),
                    "no_ruby_dy_obs": no_ruby_dy,
                    "cells": cells,
                    "linear_fit": {
                        "slope": round(slope, 4),
                        "intercept": round(intercept, 3),
                        "implied_C_if_slope_locked_to_1": round(C_implied, 3),
                    },
                }
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT_PATH}")


if __name__ == "__main__":
    main()
