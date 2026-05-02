"""b35123 p.1 per-character X coordinate measurement (Inv-X).

For each paragraph on p1, walk each character and record:
  - char
  - Information(5) X position
  - Information(6) Y position
  - prev/next char (for context)
  - is_yakumono
  - font_size

Then compare to Oxi layout JSON for same chars to find Mech 2 active
char-by-char differences.

This pinpoints which Word compresses but Oxi doesn't (or vice versa).
"""
import os
import sys
import time
import json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(
    "tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx")
OUT = os.path.abspath("pipeline_data/b35123_per_char_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")


def cls(ch):
    if ch in YAKUMONO_A: return "A"
    if ch in YAKUMONO_B: return "B"
    return "X"


def main():
    word = win32com.client.Dispatch("Word.Application")
    word.DisplayAlerts = False
    out = {"doc": DOC, "page": 1, "paragraphs": []}
    try:
        d = word.Documents.Open(DOC, ReadOnly=True)
        time.sleep(0.5)
        n_paras = d.Paragraphs.Count
        print(f"Total paragraphs: {n_paras}", flush=True)
        for pi in range(1, n_paras + 1):
            try:
                para = d.Paragraphs(pi)
                rng = para.Range
                page = int(rng.Information(3))
                if page != 1:
                    if page > 1:
                        break
                    continue
                txt = rng.Text
                if not txt or txt == "\r":
                    continue
                fs = float(rng.Font.Size) if rng.Font.Size > 0 else 12.0
                chars = rng.Characters
                p_data = {"index": pi, "text_preview": txt[:30],
                           "y": float(rng.Information(6)),
                           "font_size": fs, "chars": []}
                for ci in range(1, min(chars.Count, 80) + 1):
                    try:
                        c = chars(ci)
                        ct = c.Text
                        if ct in ("\r", "\x07"):
                            continue
                        x = float(c.Information(5))
                        y = float(c.Information(6))
                        p_data["chars"].append({
                            "ch": ct, "x": x, "y": y,
                            "cls": cls(ct),
                            "size": float(c.Font.Size),
                        })
                    except Exception:
                        continue
                out["paragraphs"].append(p_data)
                # Show progress for paragraphs with yakumono
                yak_chars = [c for c in p_data["chars"] if c["cls"] in ("A","B")]
                if yak_chars:
                    print(f"p{pi} y={p_data['y']:.0f} fs={fs} chars={len(p_data['chars'])} yak={len(yak_chars)}: {txt[:40]!r}",
                          flush=True)
            except Exception as e:
                print(f"p{pi} ERR: {e}", flush=True)
                continue
        d.Close(SaveChanges=False)
    finally:
        try: word.Quit()
        except: pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2, default=str)
    print(f"\nWrote {OUT}", flush=True)
    print(f"Total p1 paragraphs measured: {len(out['paragraphs'])}", flush=True)


if __name__ == "__main__":
    main()
