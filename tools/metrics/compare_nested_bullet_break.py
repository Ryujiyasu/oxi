"""Compare per-char x positions in Word vs Oxi for nested_bullet_08 P4."""
import win32com.client
import time
import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

DOCX = os.path.abspath('pipeline_data/docx/nested_bullet_08.docx')

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
time.sleep(0.4)

print('=== Word per-char positions for P4 (Microsoft / Azure / 連携...) ===')
para = doc.Paragraphs(4)
chars = para.Range.Characters
for ci in range(1, chars.Count + 1):
    try:
        c = chars(ci)
        ch = c.Text
        if ch in ('\r', '\x07', '\n'):
            continue
        cx = c.Information(5)  # wdHorizontalPositionRelativeToPage
        cy = c.Information(6)  # vertical pos
        line_y = round(cy, 1)
        print(f'  {ci:3d} {ch!r:6s} x={cx:6.2f} y={line_y}')
    except Exception:
        pass
doc.Close(SaveChanges=False)
word.Quit()
