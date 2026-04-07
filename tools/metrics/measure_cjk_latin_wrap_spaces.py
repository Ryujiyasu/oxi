"""COM measure space widths in cjk_latin_wrap_05.docx (the actual problematic doc)."""
import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

DOCX = os.path.abspath("pipeline_data/docx/cjk_latin_wrap_05.docx")
doc = word.Documents.Open(DOCX, ReadOnly=True)
chars = doc.Range().Characters
prev_x = None
prev_ch = None
prev_line = None
print(f"{'idx':>4s} {'ch':>3s} {'ln':>3s} {'x':>7s} {'dx':>6s}")
for ci in range(1, chars.Count + 1):
    try:
        c = chars(ci); ch = c.Text
        if ch == "\r":
            print("-- end of para --")
            continue
        if ch == "\x07":
            continue
        ln = c.Information(10)
        x = c.Information(5)
        dx = ""
        if prev_x is not None and ln == prev_line:
            dx = f"{x - prev_x:6.2f}"
        marker = " ←LB" if ln != prev_line else ""
        if ch == ' ':
            ch_disp = "SP"
        else:
            ch_disp = ch
        print(f"{ci:4d} {ch_disp:>3s} {ln:3d} {x:7.2f} {dx}{marker}")
        prev_x = x; prev_ch = ch; prev_line = ln
    except Exception:
        pass

doc.Close(False)
word.Quit()
