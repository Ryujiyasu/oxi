"""Check Word's actual char widths in style_inheritance_complex_19.

If Word applies yakumono compression, consecutive ）」 etc should show
compressed (5.5pt) advances. If not, full 11pt.
"""
import win32com.client
import os
import sys
import time

docx = os.path.abspath("pipeline_data/docx/style_inheritance_complex_19.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(docx, ReadOnly=True)
time.sleep(1)

p = doc.Paragraphs(1)
chars = p.Range.Characters
n = chars.Count
print(f"P1: {n} chars, font={p.Range.Font.Name} size={p.Range.Font.Size}")

xs = []
for ci in range(1, n + 1):
    try:
        c = chars(ci)
        ch = c.Text
        if ch in ("\r", "\x07"):
            continue
        cx = c.Information(5)
        cy = c.Information(6)
        xs.append((ch, cx, cy))
    except:
        continue

# Print width to next char
for i in range(len(xs) - 1):
    ch, x, y = xs[i]
    next_x = xs[i + 1][1]
    next_y = xs[i + 1][2]
    if abs(next_y - y) > 0.5:
        # Line break — width can't be measured
        sys.stdout.buffer.write(
            (f"  '{ch}' x={x:.2f} y={y:.2f} [LINE BREAK]\n").encode("utf-8", "replace")
        )
    else:
        w = round(next_x - x, 2)
        marker = " <-- COMPRESSED?" if w < 9 and ord(ch) >= 0x3000 else ""
        sys.stdout.buffer.write(
            (f"  '{ch}' x={x:.2f} y={y:.2f} w={w}{marker}\n").encode("utf-8", "replace")
        )

doc.Close(SaveChanges=False)
word.Quit()
