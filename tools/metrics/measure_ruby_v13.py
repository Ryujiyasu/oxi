"""V13 (base × hpsRaise × hps) focused measurement.

Reads RUBY_V13_base_{090,110,120,140}dpt.docx via COM, extracts paragraph dy,
solves for default_raise + base_ascent per base size.

Each fixture contains 9 paragraphs in a known pattern:
  P1 = no-ruby ref
  P2 = default ruby (raise=default, hps=base/2)
  P3 = no-ruby
  P4 = explicit raise=12halfpt (=6pt), hps=default
  P5 = no-ruby
  P6 = explicit raise=24halfpt (=12pt), hps=default
  P7 = no-ruby
  P8 = explicit hps=base_halfpt (=base_pt), raise=default
  P9 = no-ruby closure

dy(P_n, P_{n+1}) = P_n's line height (no-ruby) or P_n's line height + ruby_expansion (ruby).

Reference no_ruby_LH per base = base_pt × 9/7 (CLAUDE.md V10 round 4 finding).
For each ruby paragraph cell:
  measured_expansion = dy(ruby_para, next_para) − no_ruby_LH

Then derive base-aware ascent:
  expansion_explicit = max(0, raise_pt + 0.75 × hps_pt − ascent_base)
  ⇒ ascent_base = raise_pt + 0.75 × hps_pt − measured_expansion (when expansion > 0)

For default cell, default_raise is unknown — assume same ascent and solve.

Writes results to pipeline_data/ruby_v13_measurements.json.
"""
import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = os.path.abspath("pipeline_data/docx")
OUT_PATH = os.path.abspath("pipeline_data/ruby_v13_measurements.json")

V13_DOCS = [
    ("RUBY_V13_base_090dpt", 9.0),
    ("RUBY_V13_base_110dpt", 11.0),
    ("RUBY_V13_base_120dpt", 12.0),
    ("RUBY_V13_base_140dpt", 14.0),
]

CELLS = [
    # (para_idx_ruby, para_idx_next, label, explicit_raise_pt_or_None, explicit_hps_halfpt_or_None)
    (2, 3, "default",       None, None),
    (4, 5, "raise=12hp(6pt)",  6.0, None),
    (6, 7, "raise=24hp(12pt)", 12.0, None),
    (8, 9, "hps=base",         None, "base"),
]


def measure_doc(word_app, docx_path: str) -> dict:
    abs_path = os.path.abspath(docx_path)
    doc = word_app.Documents.Open(abs_path, ReadOnly=True)
    time.sleep(0.4)
    paragraphs = []
    n = doc.Paragraphs.Count
    for pi in range(1, n + 1):
        p = doc.Paragraphs(pi)
        rng = p.Range
        try:
            y = rng.Information(6)
            x = rng.Information(5)
        except Exception:
            x = y = None
        text = (rng.Text or "").replace("\r", "").replace("\x07", "")
        paragraphs.append({"i": pi, "x_pt": x, "y_pt": y, "text": text[:60]})
    doc.Close(SaveChanges=False)
    return paragraphs


def main() -> None:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    out: dict = {
        "_meta": {
            "tool": "tools/metrics/measure_ruby_v13.py",
            "purpose": "V13 base × raise × hps grid measurement",
            "cell_pattern": "P1=ref, P2=default, P3=ref, P4=raise=6pt, P5=ref, P6=raise=12pt, P7=ref, P8=hps=base, P9=ref",
        },
        "fixtures": {},
    }
    try:
        for fname, base_pt in V13_DOCS:
            print(f"\n=== {fname} (base={base_pt}pt) ===")
            docx_path = os.path.join(DOCX_DIR, fname + ".docx")
            paras = measure_doc(word, docx_path)
            no_ruby_lh = base_pt * 9.0 / 7.0
            default_hps_pt = base_pt / 2.0  # default ruby font size = base/2

            cell_results = []
            for ruby_idx, next_idx, label, raise_pt, hps_marker in CELLS:
                if next_idx > len(paras):
                    print(f"  cell {label} -> SKIP (paragraph index out of range)")
                    continue
                ruby_para = paras[ruby_idx - 1]
                next_para = paras[next_idx - 1]
                if ruby_para["y_pt"] is None or next_para["y_pt"] is None:
                    print(f"  cell {label} -> SKIP (None y)")
                    continue
                dy = next_para["y_pt"] - ruby_para["y_pt"]
                expansion = dy - no_ruby_lh
                hps_pt_used = base_pt if hps_marker == "base" else default_hps_pt
                cell_results.append({
                    "label": label,
                    "ruby_para_idx": ruby_idx,
                    "next_para_idx": next_idx,
                    "ruby_y_pt": round(ruby_para["y_pt"], 3),
                    "next_y_pt": round(next_para["y_pt"], 3),
                    "dy_pt": round(dy, 3),
                    "no_ruby_lh_pt": round(no_ruby_lh, 3),
                    "expansion_pt": round(expansion, 3),
                    "explicit_raise_pt": raise_pt,
                    "hps_pt": hps_pt_used,
                })
                print(
                    f"  cell {label}: dy={dy:.3f} - no_ruby_lh={no_ruby_lh:.3f} "
                    f"= expansion={expansion:.3f}pt (raise={raise_pt} hps={hps_pt_used})"
                )

            # Reference dy from neighboring no-ruby pair (P1 -> P2 wraps ruby para,
            # but P3 -> P4 wraps no-ruby->ruby; P9 has no successor).
            # Use dy(P1,P2)? Actually P1 is no-ruby and P2 is the first ruby — dy(P1,P2) = P1's height (no-ruby).
            no_ruby_observed = []
            for (a, b) in [(1, 2), (3, 4), (5, 6), (7, 8)]:
                if a <= len(paras) and b <= len(paras):
                    pa = paras[a-1]
                    pb = paras[b-1]
                    if pa["y_pt"] is not None and pb["y_pt"] is not None:
                        no_ruby_observed.append(round(pb["y_pt"] - pa["y_pt"], 3))

            out["fixtures"][fname] = {
                "base_pt": base_pt,
                "default_hps_pt": default_hps_pt,
                "no_ruby_lh_predicted_pt": round(no_ruby_lh, 3),
                "no_ruby_lh_observed_pt": no_ruby_observed,
                "all_paragraphs": [
                    {"i": p["i"], "y_pt": round(p["y_pt"], 3) if p["y_pt"] is not None else None, "text": p["text"]}
                    for p in paras
                ],
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
