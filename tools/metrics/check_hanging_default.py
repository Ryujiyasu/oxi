"""Check HangingPunctuation default in new docs vs opened docs."""
import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

# New doc
doc = word.Documents.Add()
doc.Range().InsertAfter("テスト")
para = doc.Paragraphs(1)
print(f"NEW DOC: HangingPunct={para.Format.HangingPunctuation} FELB={para.Format.FarEastLineBreakControl}")
doc.Close(False)

# Open existing doc
for d in ["ruby_text_lineheight_11", "special_chars_spacing_01"]:
    p = os.path.abspath(f"pipeline_data/docx/{d}.docx")
    doc = word.Documents.Open(p, ReadOnly=True)
    para = doc.Paragraphs(1)
    print(f"{d}: HangingPunct={para.Format.HangingPunctuation} FELB={para.Format.FarEastLineBreakControl}")
    doc.Close(False)

word.Quit()
