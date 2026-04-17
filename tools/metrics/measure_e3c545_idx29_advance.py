"""COM char-by-char advance for e3c545 idx=29 to quantify char-advance bug.

Memory: Oxi wraps idx=29 to 44+2 chars, Word fits all 46. 20.5pt drift
cascade from this. Per project_yakumono_advance_fix_location.md, Word
caps yakumono at 10.5pt regardless of font size.

This script measures Word's per-char x positions for each char in
paras[29] and computes advance(i) = x[i+1] - x[i]. Then compares
against Oxi's computed advance from char_width_pt().
"""
import json, time, sys
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path(r"C:/Users/ryuji/oxi-4/tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx")
OUT = Path(__file__).with_name("output") / "e3c545_idx29_advance.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

word = win32com.client.Dispatch("Word.Application")
time.sleep(1.0)
word.Visible = False
word.DisplayAlerts = False

try:
    doc = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
    time.sleep(1.5)

    # paras[29] is 0-based XML. COM is 1-based. XML paras[29] = COM Paragraphs(30).
    target_idx = 30
    p = doc.Paragraphs(target_idx)
    rng = p.Range
    text = rng.Text
    print(f"Paragraph {target_idx}: {len(text)} chars, text[:80]={text[:80]!r}")

    sel = word.Selection
    chars = []
    for ci in range(rng.Start, rng.End):
        sel.SetRange(ci, ci + 1)
        try:
            y = sel.Information(6)
            x = sel.Information(5)
            pg = int(sel.Information(3))
            ch = sel.Text
        except Exception:
            continue
        chars.append({
            "ci": ci - rng.Start,
            "page": pg,
            "y": round(y, 2),
            "x": round(x, 2),
            "ch": ch,
            "codepoint": ord(ch) if len(ch) == 1 else None,
        })

    doc.Close(False)

    # Group by line (y)
    from collections import defaultdict
    lines = defaultdict(list)
    for c in chars:
        lines[(c['page'], c['y'])].append(c)

    print(f"\n{len(chars)} chars, {len(lines)} lines total:")
    for (pg, y) in sorted(lines.keys()):
        chars_in_line = sorted(lines[(pg, y)], key=lambda c: c['x'])
        text_joined = ''.join(c['ch'] for c in chars_in_line)
        # Compute advances
        print(f"\n  --- page {pg} y={y:.1f} — {len(chars_in_line)} chars ---")
        print(f"  text: {text_joined[:80]!r}")
        print(f"  per-char advance (x[i+1] - x[i]):")
        for i in range(min(len(chars_in_line) - 1, 50)):
            c = chars_in_line[i]
            nc = chars_in_line[i + 1]
            adv = round(nc['x'] - c['x'], 2)
            cp = c.get('codepoint')
            cp_str = f"U+{cp:04X}" if cp else "?"
            print(f"    [{i:2d}] ch={c['ch']!r:<6} {cp_str}  x={c['x']:>6.2f}  next_x={nc['x']:>6.2f}  adv={adv}pt")

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(chars, f, ensure_ascii=False, indent=2)
    print(f"\nSaved → {OUT}")

finally:
    try: word.Quit()
    except: pass
