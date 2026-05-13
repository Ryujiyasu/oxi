"""Measure exact char widths for e3c545 paragraph idx=29 (1-indexed 30).

The paragraph text is:
  メタデータは、各機関で独自に定義します。具体例は、「９．例 (1)メタデータ」を参照ください。

Day 34 part 22 hypothesis: Oxi wraps to 2 lines (drift origin -20.5pt),
Word fits 1 line via yakumono compression of 6 pair occurrences
(、×2, 。×2, 「, 」).

This script verifies by measuring each char's X advance via Word COM.
"""
import os
import sys
import time
import json
import win32com.client

DOCX = os.path.abspath(r"tools\golden-test\documents\docx\e3c545fac7a7_LOD_Handbook.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
time.sleep(2)

# Get paragraph at 1-indexed position 30 (= 0-indexed 29)
p = doc.Paragraphs(30)
rng = p.Range
print(f"Para text: {rng.Text!r}")
print(f"  Font: {rng.Font.Name} {rng.Font.Size}pt")

# Get start position for context
start_info = doc.Range(rng.Start, rng.Start).Information(6)  # vert pos
print(f"  Start Y: {start_info}")

chars = rng.Characters
n = chars.Count
print(f"  n chars: {n}")

records = []
for i in range(1, n + 1):
    try:
        c = chars(i)
        ch = c.Text
        if ch in ("\r", "\x07", "\n"):
            continue
        cx = c.Information(5)  # horiz pos
        cy = c.Information(6)  # vert pos
        records.append({"i": i, "ch": ch, "x": cx, "y": cy, "ord": ord(ch[0]) if ch else 0})
    except Exception as e:
        print(f"  err at {i}: {e}")

# Compute advances and identify line structure
print()
print("=== Char-by-char advances ===")
prev_y = None
line_idx = 0
for k in range(len(records)):
    r = records[k]
    nxt = records[k + 1] if k + 1 < len(records) else None
    if prev_y is None or abs(r["y"] - prev_y) > 1.0:
        if prev_y is not None:
            line_idx += 1
        print(f"--- line {line_idx} (y={r['y']:.2f}) ---")
        prev_y = r["y"]
    if nxt and abs(nxt["y"] - r["y"]) < 1.0:
        adv = round(nxt["x"] - r["x"], 2)
    else:
        adv = None
    # Mark yakumono chars
    yakumono = r["ord"] in (0x3001, 0x3002, 0x300C, 0x300D, 0xFF0C, 0xFF0E, 0xFF08, 0xFF09)
    marker = "  ← YAKU" if yakumono else ""
    is_fullwidth = (0x3000 <= r["ord"] <= 0x9FFF) or (0xFF00 <= r["ord"] <= 0xFFEF)
    print(f"  i={r['i']:3} ch={r['ch']!r} x={r['x']:7.2f} y={r['y']:6.2f} adv={adv}{marker} fw={is_fullwidth}")

# Compute summary by line
print()
print("=== Summary ===")
lines = {}
for r in records:
    y = round(r["y"], 1)
    lines.setdefault(y, []).append(r)
for y in sorted(lines.keys()):
    chars_in_line = lines[y]
    if len(chars_in_line) < 2:
        continue
    line_w = chars_in_line[-1]["x"] - chars_in_line[0]["x"]
    n_yaku = sum(1 for r in chars_in_line if r["ord"] in (0x3001, 0x3002, 0x300C, 0x300D))
    print(f"  y={y:.1f}: {len(chars_in_line)} chars, line span={line_w:.2f}pt, yakumono count={n_yaku}")
    print(f"    text: {''.join(r['ch'] for r in chars_in_line)!r}")

doc.Close(False)
word.Quit()

out_path = "pipeline_data/e3c545_idx29_widths.json"
with open(out_path, "w", encoding="utf-8") as f:
    json.dump({"records": records, "lines": {str(y): [r for r in chars_in_line] for y, chars_in_line in lines.items()}}, f, ensure_ascii=False, indent=2, default=str)
print(f"\nSaved: {out_path}")
