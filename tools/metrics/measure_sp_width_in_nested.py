"""Measure exact space widths around CJK in nested_bullet_08 vs short doc."""
import win32com.client
import time
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

import sys as _s
DOC_NAME = _s.argv[1] if len(_s.argv) > 1 else 'nested_bullet_08'
PARA = int(_s.argv[2]) if len(_s.argv) > 2 else 4
DOC = os.path.abspath(f'pipeline_data/docx/{DOC_NAME}.docx')

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(0.5)

para = doc.Paragraphs(PARA)
chars = para.Range.Characters
xs = []
for ci in range(1, chars.Count + 1):
    try:
        c = chars(ci)
        ch = c.Text
        if ch in ('\r', '\x07', '\n'):
            continue
        cx = c.Information(5)
        cy = c.Information(6)
        xs.append((ci, ch, cx, cy))
    except Exception:
        pass

# Compute per-char width = next.x - this.x (for chars on same line)
print(f'{"#":<4}{"ch":<6}{"x":>8}{"width":>8}{"context":<30}')
for i in range(len(xs) - 1):
    idx, ch, x, y = xs[i]
    next_idx, next_ch, next_x, next_y = xs[i + 1]
    if next_y != y:
        continue  # skip line breaks
    w = next_x - x
    ctx = ''
    if ch == ' ':
        prev_ch = xs[i - 1][1] if i > 0 else ''
        ctx = f'{prev_ch}<sp>{next_ch}'
    print(f'{idx:<4}{ch!r:<6}{x:>8.2f}{w:>8.2f}  {ctx}')

doc.Close(SaveChanges=False)
word.Quit()
