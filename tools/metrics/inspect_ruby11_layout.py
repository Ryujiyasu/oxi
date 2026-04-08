"""Open ruby_text_lineheight_11.docx and extract per-char layout from Word.

Goal: understand how Word fits 40 chars/line in a doNotCompress doc where
naive 11pt × 40 = 440pt > 432pt available width.

Hypotheses:
1. Burasagari (yakumono hangs past right margin)
2. Width calculation differs (some narrowing applied even in doNotCompress mode)
3. Char width is something other than 11pt
"""
import win32com.client
import time
import os
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

path = os.path.abspath("pipeline_data/docx/ruby_text_lineheight_11.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(0.4)

ps = doc.PageSetup
print(f"PageWidth={ps.PageWidth:.2f}pt L={ps.LeftMargin:.2f}pt R={ps.RightMargin:.2f}pt")
print(f"  body width = {ps.PageWidth - ps.LeftMargin - ps.RightMargin:.2f}pt")
print()

# Per-char info, all paragraphs
for pi in [1, 2, 3]:
    try:
        para = doc.Paragraphs(pi)
    except Exception:
        break
    text = para.Range.Text
    if len(text) < 3:
        continue
    print(f"=== Paragraph {pi}: {text[:40]!r}... (len={len(text)}) ===")
    chars = para.Range.Characters
    rows = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            cx = c.Information(5)
            cy = c.Information(6)
            font_name = c.Font.Name
            font_size = c.Font.Size
            rows.append((ci, ch, cx, cy, font_name, font_size))
        except Exception:
            continue
    # Group by line
    lines = {}
    for r in rows:
        lines.setdefault(round(r[3], 1), []).append(r)
    for y in sorted(lines.keys()):
        ln = lines[y]
        first_x = min(r[2] for r in ln)
        last_x = max(r[2] for r in ln)
        text_line = "".join(r[1] for r in ln)
        font = ln[0][4]
        size = ln[0][5]
        print(f"  y={y} chars={len(ln)} first_x={first_x:.2f} last_x={last_x:.2f} width={last_x-first_x:.2f}pt font={font!r} size={size}")
        print(f"    text: {text_line!r}")
    print()

doc.Close(SaveChanges=False)
word.Quit()
