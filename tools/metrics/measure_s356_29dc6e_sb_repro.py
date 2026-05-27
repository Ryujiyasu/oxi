"""S356 — COM-measure each minimal repro to extract Word's actual sb behavior.

For each (signature, sb_value) docx, measure:
  row1_y = first paragraph's y (first-char position via Information(6))
  row2_y = second paragraph's y (which is row 2 first-in-cell paragraph)
  delta = row2_y - row1_y

Then for each signature, sb_applied = delta(B) - delta(A).
  - sb_applied ≈ 7.3pt → Word applies sb (S136 rule correct here)
  - sb_applied ≈ 0pt   → Word suppresses sb
"""
import json
import sys
from pathlib import Path
import win32com.client

REPRO_DIR = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\repros\s356_29dc6e_sb").resolve()
OUT_PATH = Path(r"c:\Users\ryuji\oxi-main\tools\metrics\s356_29dc6e_sb_word.json").resolve()

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = False

results = {}

try:
    for docx in sorted(REPRO_DIR.glob("*.docx")):
        try:
            doc = word.Documents.Open(str(docx), ReadOnly=True)
            doc.ActiveWindow.View.Type = 3
            doc.Repaginate()

            # Get table cell paragraphs (doc.Paragraphs includes body+table)
            # Find ROW1_REF and ROW2_TARGET by text content
            target_p = {"ROW1_REF": None, "ROW2_TARGET": None}
            n_paras = doc.Paragraphs.Count
            for i in range(1, n_paras + 1):
                p = doc.Paragraphs(i)
                text = (p.Range.Text or "").rstrip("\r\x07").strip()
                if text in target_p and target_p[text] is None:
                    target_p[text] = p
            if target_p["ROW1_REF"] is None or target_p["ROW2_TARGET"] is None:
                results[docx.stem] = {
                    "error": "could not find ROW1_REF or ROW2_TARGET",
                    "n_paras": n_paras,
                }
                doc.Close(SaveChanges=False)
                continue

            r1 = target_p["ROW1_REF"]
            r2 = target_p["ROW2_TARGET"]
            r1_rng = r1.Range
            r2_rng = r2.Range
            r1_start = doc.Range(r1_rng.Start, r1_rng.Start)
            r2_start = doc.Range(r2_rng.Start, r2_rng.Start)

            r1_y = r1_start.Information(6)
            r2_y = r2_start.Information(6)

            results[docx.stem] = {
                "n_paras": n_paras,
                "row1_text": r1_rng.Text.rstrip("\r\x07")[:30],
                "row2_text": r2_rng.Text.rstrip("\r\x07")[:30],
                "row1_y": r1_y,
                "row2_y": r2_y,
                "delta_y": r2_y - r1_y,
                "row2_SpaceBefore_pt": r2.Format.SpaceBefore,
                "row2_LineSpacing_pt": r2.Format.LineSpacing,
                "row2_LineSpacingRule": r2.Format.LineSpacingRule,
            }
            doc.Close(SaveChanges=False)
        except Exception as e:
            results[docx.stem] = {"error": str(e)}
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass

    OUT_PATH.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote: {OUT_PATH}")

    # Summarize: pair each signature's A vs B
    print(f"\n{'signature':<25} {'Δy(A)':>7} {'Δy(B)':>7} {'sb_applied':>11} {'verdict':>15}")
    signatures = ["E0_before_only", "E1_before_lineRule", "E2_before_lineRule_bl",
                  "E3_full_29dc6e_no_grid", "E4_full_29dc6e_w_grid"]
    for sig in signatures:
        a = results.get(f"{sig}_A_sb0", {})
        b = results.get(f"{sig}_B_sb146", {})
        if "delta_y" in a and "delta_y" in b:
            sb_applied = b["delta_y"] - a["delta_y"]
            if abs(sb_applied) < 0.5:
                verdict = "SUPPRESSED"
            elif abs(sb_applied - 7.3) < 0.5:
                verdict = "APPLIED ~7.3pt"
            else:
                verdict = f"OTHER {sb_applied:+.2f}"
            print(f"{sig:<25} {a['delta_y']:>7.2f} {b['delta_y']:>7.2f} {sb_applied:>+11.2f} {verdict:>15}")
        else:
            print(f"{sig:<25} {'?':>7} {'?':>7} {'?':>11} {'ERROR':>15}")

finally:
    word.Quit()
