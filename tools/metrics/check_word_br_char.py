"""Check what char COM returns for <w:br/> soft line break."""
import win32com.client
import os
import sys
import time

docx = os.path.abspath("pipeline_data/docx/bullet_list_japanese_01.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(docx, ReadOnly=True)
time.sleep(1)

p = doc.Paragraphs(1)
chars = p.Range.Characters
print(f"P1 has {chars.Count} chars")
for ci in range(1, min(chars.Count + 1, 80)):
    try:
        c = chars(ci)
        ch = c.Text
        sys.stdout.buffer.write(
            (f"  {ci:3d}: U+{ord(ch):04X} {repr(ch)}\n").encode("utf-8", "replace")
        )
    except:
        continue

doc.Close(SaveChanges=False)
word.Quit()
