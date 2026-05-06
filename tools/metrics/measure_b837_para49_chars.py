"""b837 p4 para 49 per-character X/Y measurement (Day 1, Session 55+).

Goal: ground-truth Bug B grid_char_pitch over-pack on b837. Para 49 in Oxi
fits all 37 chars on 1 line (last char ends at x=548, page right margin=524).
Word renders it as 2 lines per visual inspection. This measures Word's exact
wrap point and per-char x positions.

Output: pipeline_data/b837_para49_word_chars.json
"""
import os
import sys
import time
import json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(
    "tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
OUT = os.path.abspath("pipeline_data/b837_para49_word_chars.json")

# We measure paragraphs near 49 (Oxi 0-based = Word 1-based 50)
# Want to see: para 49 and 50 (which Oxi rendered short)
# Word indices (1-based): 50, 51 (for Oxi indices 49, 50)
TARGETS = [10, 11, 14, 17, 30, 49, 55, 56]  # 0-based, Word_i = +1


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.DisplayAlerts = False
    out = {"doc": DOC, "paragraphs": []}
    try:
        d = word.Documents.Open(DOC, ReadOnly=True)
        time.sleep(0.5)
        for oxi_idx in TARGETS:
            wi = oxi_idx + 1  # word 1-based
            try:
                para = d.Paragraphs(wi)
                rng = para.Range
                start_rng = d.Range(rng.Start, rng.Start)
                page = int(start_rng.Information(3))
                start_y = float(start_rng.Information(6))
                start_x = float(start_rng.Information(5))
                txt = rng.Text
                fs = float(rng.Font.Size) if rng.Font.Size > 0 else 12.0
                chars = rng.Characters
                p_data = {
                    "oxi_idx": oxi_idx,
                    "word_i": wi,
                    "text_preview": txt[:60],
                    "page": page,
                    "start_y": start_y,
                    "start_x": start_x,
                    "font_size": fs,
                    "chars": [],
                }
                # Walk every character
                for ci in range(1, min(chars.Count + 1, 80)):
                    try:
                        c = chars(ci)
                        ct = c.Text
                        if ct in ("\r", "\x07"):
                            continue
                        # Use collapsed start for this char too
                        crng = c
                        cx = float(crng.Information(5))
                        cy = float(crng.Information(6))
                        cpage = int(crng.Information(3))
                        p_data["chars"].append({
                            "i": ci,
                            "ch": ct,
                            "x": cx,
                            "y": cy,
                            "page": cpage,
                            "size": float(c.Font.Size) if c.Font.Size > 0 else fs,
                        })
                    except Exception as e:
                        p_data["chars"].append({"i": ci, "error": str(e)[:80]})
                out["paragraphs"].append(p_data)
                # Print summary
                ys = sorted({c["y"] for c in p_data["chars"] if "y" in c})
                print(f"Para {oxi_idx} (word_i={wi}): page={page} start=({start_x:.1f},{start_y:.1f}) "
                      f"text={txt[:30]!r} #lines={len(ys)} y_values={ys[:4]}", flush=True)
            except Exception as e:
                print(f"  Para {oxi_idx} ERROR: {e}", flush=True)
        d.Close(SaveChanges=False)
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"\nSaved: {OUT}")


if __name__ == "__main__":
    main()
