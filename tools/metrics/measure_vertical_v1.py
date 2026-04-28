"""VW_V1 — measure Word's tbRlV cell rendering geometry via COM.

For each VW_V1 fixture (single-cell tbRlV), open in Word and measure:
  - Cell bounding box (x, y, width, height) via Range.Information(5)/(6)/(7)/(8)
  - Per-character X/Y in the vertical text (Range.Characters iteration)
  - Whether characters are rotated (Font properties may indicate)

The expected behavior (per ECMA-376 §17.18.93 textDirection):
  tbRlV: top-to-bottom writing, right-to-left line progression, vertical glyphs
  - First char at top of cell content area
  - Subsequent chars stack downward
  - When column overflows cell height, new column to the LEFT
  - CJK chars are rendered upright (vertical glyph variants used)
  - Latin chars are usually rotated 90° CW

Writes pipeline_data/vertical_v1_measurements.json.
"""
import json
import os
import sys
import time

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = os.path.abspath("pipeline_data/docx")
OUT_PATH = os.path.abspath("pipeline_data/vertical_v1_measurements.json")

VW_V1_FIXTURES = [
    "VW_V1_basic",
    "VW_V1_long",
    "VW_V1_msmincho_14pt",
    "VW_V1_yu_mincho",
    "VW_V1_two_cols",
]


def measure_doc(word_app, docx_path: str) -> dict:
    abs_path = os.path.abspath(docx_path)
    doc = word_app.Documents.Open(abs_path, ReadOnly=True)
    time.sleep(0.4)
    out: dict = {"paragraphs": []}
    n_paras = doc.Paragraphs.Count
    for pi in range(1, n_paras + 1):
        p = doc.Paragraphs(pi)
        rng = p.Range
        try:
            y = rng.Information(6)  # vertical position relative to page
            x = rng.Information(5)  # horizontal position relative to page
            page = rng.Information(3)
        except Exception:
            x = y = page = None
        text = (rng.Text or "").replace("\r", "").replace("\x07", "")
        para_data = {
            "i": pi,
            "page": page,
            "x_pt": x,
            "y_pt": y,
            "text": text[:60],
            "chars": [],
        }
        # Per-char positions
        try:
            chars = rng.Characters
            cnt = chars.Count
            for ci in range(1, cnt + 1):
                c = chars(ci)
                ch = c.Text
                if ch in ("\r", "\x07"):
                    continue
                try:
                    cx = c.Information(5)
                    cy = c.Information(6)
                    font = c.Font.Name
                    sz = c.Font.Size
                    para_data["chars"].append({
                        "i": ci, "ch": ch,
                        "x_pt": round(cx, 3), "y_pt": round(cy, 3),
                        "font": font, "sz": sz,
                    })
                except Exception:
                    pass
        except Exception:
            pass
        out["paragraphs"].append(para_data)
    doc.Close(SaveChanges=False)
    return out


def main() -> None:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out: dict = {
        "_meta": {
            "tool": "tools/metrics/measure_vertical_v1.py",
            "purpose": "measure Word tbRlV cell-level vertical writing per-char X/Y",
        },
        "fixtures": {},
    }
    try:
        for fname in VW_V1_FIXTURES:
            print(f"\n=== {fname} ===")
            docx_path = os.path.join(DOCX_DIR, fname + ".docx")
            data = measure_doc(word, docx_path)
            out["fixtures"][fname] = data
            for p in data["paragraphs"]:
                print(f"  P{p['i']} page={p['page']} x={p['x_pt']:.2f} y={p['y_pt']:.2f}  text={p['text']!r}")
                for c in p["chars"][:8]:
                    print(f"    [{c['i']:>2}] {c['ch']!r:6} x={c['x_pt']:7.3f} y={c['y_pt']:7.3f} font={c['font']:<12} sz={c['sz']}")
                if len(p["chars"]) > 8:
                    print(f"    ... ({len(p['chars']) - 8} more)")
    finally:
        word.Quit()

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {OUT_PATH}")


if __name__ == "__main__":
    main()
