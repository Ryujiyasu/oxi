"""Re-save jfmb via Word COM to capture Word's normalization pass.

Open the on-disk jfmb, save as new path, then we can diff document.xml
to identify what Word adds that triggers CJK-adjacent space widening.
"""
import win32com.client
import os
import time
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = os.path.abspath("pipeline_data/docx/japanese_font_mixing_baseline.docx")
DST = os.path.abspath(
    "pipeline_data/docx/japanese_font_mixing_baseline_runtime.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
try:
    d = word.Documents.Open(SRC, ReadOnly=False)
    time.sleep(0.5)
    d.SaveAs2(DST, FileFormat=12)
    d.Close(SaveChanges=False)
    print(f"Saved: {DST}")
finally:
    word.Quit()
