"""COM-measure per-char x positions in special_chars_spacing_01.docx
to determine actual char widths Word uses (vs Oxi's 11pt per char)."""
import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

DOCX = os.path.abspath("pipeline_data/docx/special_chars_spacing_01.docx")
doc = word.Documents.Open(DOCX, ReadOnly=True)
chars = doc.Range().Characters
prev_x = None
prev_ch = None
prev_line = None
print(f"{'idx':>4s}  {'ch':>3s}  {'line':>4s}  {'x':>7s}  {'dx':>6s}")
for ci in range(1, chars.Count + 1):
    try:
        c = chars(ci)
        ch = c.Text
        if ch in ("\r", "\x07"):
            continue
        ln = c.Information(10)  # wdFirstCharacterLineNumber
        x = c.Information(5)
        dx = ""
        if prev_x is not None and ln == prev_line:
            dx = f"{x - prev_x:6.2f}"
        marker = "" if ln == prev_line else " ← LB"
        print(f"{ci:4d}  {ch}  {ln:4d}  {x:7.2f}  {dx}{marker}")
        prev_x = x
        prev_ch = ch
        prev_line = ln
    except Exception:
        pass

doc.Close(SaveChanges=False)
word.Quit()
