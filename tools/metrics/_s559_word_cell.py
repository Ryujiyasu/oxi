# -*- coding: utf-8 -*-
"""S559 — COM-measure the REAL 3a4f ⑦ cell (para 2234) in Word: per-char x
(Information(5)) + y (Information(6)) to find where Word wraps (line-1 char
count + right edge) and the cell-left. Compares to Oxi (1 line, x0=113.0,
right 522.5, 39 chars). Determines the true cell content width → which Oxi
cell bug (firstLine shift / cellMar / autofit width) is the −1-char source.
"""
import os
import sys

import win32com.client as w32

DOCX = r'c:\tmp\3a4f9f.docx'
ANCHOR = u'常に整理整頓に努め'

word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    wdoc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
    try:
        rng = wdoc.Content
        f = rng.Find
        f.Text = ANCHOR
        if not f.Execute():
            print('anchor not found')
            sys.exit(1)
        # rng now = the found anchor; expand to its paragraph
        para = rng.Paragraphs(1).Range
        start = para.Start
        end = para.End
        sys.stdout.reconfigure(encoding='utf-8')
        rows = []
        for i in range(min(end - start, 60)):
            r = wdoc.Range(start + i, start + i + 1)
            ch = r.Text or ''
            if ch in ('\r', '\n', '\x07', '\x0b'):
                continue
            x = wdoc.Range(start + i, start + i).Information(5)
            y = wdoc.Range(start + i, start + i).Information(6)
            rows.append((ch, x, y))
        # group into lines
        lines = []
        cur = []
        y0 = None
        for ch, x, y in rows:
            if y0 is None:
                y0 = y
            if abs(y - y0) > 0.5:
                lines.append(cur)
                cur = []
                y0 = y
            cur.append((ch, x))
        if cur:
            lines.append(cur)
        print('WORD ⑦ cell: %d lines' % len(lines))
        for li, ln in enumerate(lines):
            x0 = ln[0][1]
            xlast = ln[-1][1]
            print('  L%d n=%d  x0=%.2f  last_char_x=%.2f  text=%s'
                  % (li + 1, len(ln), x0, xlast, ''.join(c for c, _ in ln)))
        # line-1 right edge estimate: last char x + ~10.5 (fullwidth) or measure next line gap
        if lines:
            l1 = lines[0]
            print('  L1 last char=%r at x=%.2f (right edge ~= +adv)' % (l1[-1][0], l1[-1][1]))
        # also the cell width: Information(5) of the para start, and the table cell
        print('  para start x = %.2f' % wdoc.Range(start, start).Information(5))
    finally:
        wdoc.Close(False)
finally:
    word.Quit()
