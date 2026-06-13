# -*- coding: utf-8 -*-
"""S557c — quantify the L3 under-pack. Dump Word per-char advances for the
isolated para, sum L3 (Word 40 chars) and compare to the justified content
width W (taken from a full matched line, e.g. L4=41 in both). Then sum what
Oxi's 38-char L3 + the 39th char (化) would need, and show the deficit and
which compressions (openers / digits) close it.
"""
import json
import os
import sys

import win32com.client as w32

DOCX = os.path.abspath('tools/golden-test/repros/s557_isolate/s557_none.docx')

word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    wdoc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        pr = wdoc.Paragraphs(1).Range
        start = pr.Start
        n = pr.End - pr.Start
        rows = []
        for i in range(min(n, 420)):
            rng = wdoc.Range(start + i, start + i + 1)
            ch = rng.Text or ''
            if ch in ('\r', '\x07', '\n', '\x0b'):
                continue
            x = wdoc.Range(start + i, start + i).Information(5)
            y = wdoc.Range(start + i, start + i).Information(6)
            rows.append((ch, x, y))
        # line right edge: maximum x on a line + last char advance is unknown,
        # so estimate the line END as the x of the FIRST char of the next line's
        # left margin == constant left x. The justified RIGHT edge = left_x +
        # content_width. Compute content_width from a full line as (last char x +
        # its natural advance). Simpler: take per-line span = max_x - min_x; the
        # full justified line with the largest span ~ content width.
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
        sys.stdout.reconfigure(encoding='utf-8')
        left_x = min(x for ln in lines for _, x in ln)
        print('left_x = %.3f' % left_x)
        for li, ln in enumerate(lines):
            span = ln[-1][1] - ln[0][1]
            print('L%d n=%d  first_x=%.2f last_x=%.2f span(excl last adv)=%.2f'
                  % (li + 1, len(ln), ln[0][1], ln[-1][1], span))
        # L3 detail: sum advances (all but last, which is unknown)
        l3 = lines[2]
        print('--- L3 advances ---')
        tot = 0.0
        for k in range(len(l3) - 1):
            ch, x = l3[k]
            adv = l3[k + 1][1] - x
            tot += adv
            print('  %s %.2f' % (ch, adv))
        print('L3 sum(first %d chars) = %.2f, last char=%s' % (len(l3) - 1, tot, l3[-1][0]))
    finally:
        wdoc.Close(False)
finally:
    word.Quit()
