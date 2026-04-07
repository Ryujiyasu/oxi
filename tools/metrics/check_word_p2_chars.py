"""Check char-by-char what Word puts on each line for paragraph_spacing_grid_04 P2."""
import win32com.client
import os
import sys
import time

docx = os.path.abspath("pipeline_data/docx/paragraph_spacing_grid_04.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(docx, ReadOnly=True)
time.sleep(1)

# Para 2 (the one with the dch issue)
p = doc.Paragraphs(2)
chars = p.Range.Characters
print(f"P2 has {chars.Count} chars")

prev_y = None
line_chars = []
line_no = 0
for ci in range(1, chars.Count + 1):
    try:
        c = chars(ci)
        ch = c.Text
        if ch in ("\r", "\x07"):
            continue
        cy = c.Information(6)
        if prev_y is not None and abs(cy - prev_y) > 0.5:
            sys.stdout.buffer.write(
                (f"  L{line_no} ({len(line_chars)}ch): " + "".join(line_chars) + "\n").encode("utf-8", "replace")
            )
            line_no += 1
            line_chars = []
        prev_y = cy
        line_chars.append(ch)
    except:
        continue
if line_chars:
    sys.stdout.buffer.write(
        (f"  L{line_no} ({len(line_chars)}ch): " + "".join(line_chars) + "\n").encode("utf-8", "replace")
    )

doc.Close(SaveChanges=False)
word.Quit()
