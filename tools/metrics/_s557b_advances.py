# -*- coding: utf-8 -*-
"""S557b — per-char x-advance on the isolated d77a para9 repro, grouped by
line (Information(6) Y to detect line breaks, Information(5) X for advance).
Shows HOW Word fits 40 chars on L3 (bracket pair-halving) vs Oxi's 38.
Usage: python _s557b_advances.py
"""
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
        # collect (char, x, y) per non-control char
        rows = []
        for i in range(min(n, 420)):
            rng = wdoc.Range(start + i, start + i + 1)
            ch = rng.Text or ''
            if ch in ('\r', '\x07', '\n', '\x0b'):
                continue
            x = wdoc.Range(start + i, start + i).Information(5)
            y = wdoc.Range(start + i, start + i).Information(6)
            rows.append((ch, x, y))
        # group into lines by y
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
        for li, ln in enumerate(lines):
            # advance = next x - this x (last char advance from line width unknown -> mark)
            print('--- L%d (n=%d) ---' % (li + 1, len(ln)))
            parts = []
            for k in range(len(ln)):
                ch, x = ln[k]
                if k + 1 < len(ln):
                    adv = ln[k + 1][1] - x
                else:
                    adv = float('nan')
                parts.append('%s%.2f' % (ch, adv))
            print(' '.join(parts))
    finally:
        wdoc.Close(False)
finally:
    word.Quit()
