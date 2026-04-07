"""Check what font Word actually uses for half-width parens in CJK context."""
import win32com.client
import os
import sys

docx = os.path.abspath("pipeline_data/docx/ruby_text_01.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(docx, ReadOnly=True)

import time
time.sleep(1)

# First paragraph chars
p = doc.Paragraphs(1)
chars = p.Range.Characters
for ci in range(1, min(chars.Count + 1, 20)):
    c = chars(ci)
    ch = c.Text
    if ch in ("\r", "\x07"):
        continue
    try:
        font_name = c.Font.Name
        font_size = c.Font.Size
        # Get vertical position
        y = c.Information(6)
        x = c.Information(5)
        ascii_or_ea = "ASCII" if ord(ch) < 128 else "EA"
        sys.stdout.buffer.write(
            (f"  '{ch}' ({ascii_or_ea}): font={font_name} size={font_size} y={y:.3f} x={x:.3f}\n").encode("utf-8", errors="replace")
        )
    except Exception as e:
        print(f"  err {ch}: {e}")

doc.Close(SaveChanges=False)
word.Quit()
