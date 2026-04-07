"""Open actual ruby_text_lineheight_11.docx and dump per-char line numbers
to verify Word's resolution strategy on the real document (which has
w:characterSpacingControl=doNotCompress)."""
import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

DOCX = os.path.abspath("pipeline_data/docx/ruby_text_lineheight_11.docx")
WD_FIRST_LINE = 10

doc = word.Documents.Open(DOCX, ReadOnly=True)
chars = doc.Range().Characters
print(f"Total chars: {chars.Count}")
print(f"{'idx':>4s}  {'ch':>3s}  {'line':>4s}  {'x':>7s}")
print("-" * 32)
prev_line = 0
para_idx = 1
for ci in range(1, chars.Count + 1):
    try:
        c = chars(ci)
        ch = c.Text
        if ch == "\r":
            print(f"  -- end of para {para_idx} --")
            para_idx += 1
            continue
        if ch == "\x07":
            continue
        ln = c.Information(WD_FIRST_LINE)
        x = c.Information(5)  # wdHorizontalPositionRelativeToPage
        marker = ""
        if ln != prev_line:
            marker = " ← line start"
            prev_line = ln
        print(f"{ci:4d}  {ch}  {ln:4d}  {x:7.2f}{marker}")
    except Exception as e:
        print(f"  err at {ci}: {e}")

doc.Close(SaveChanges=False)
word.Quit()
