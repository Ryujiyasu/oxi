# -*- coding: utf-8 -*-
"""Measure Word's per-char advance on ikujidetail body para i=199 (1-based),
which Oxi over-wraps 1->2 lines. Information(5)=wdHorizontalPositionRelToPage.
Determines whether the no-type docGrid charSpace=-3531 compresses chars."""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
DOCX = os.path.abspath('tools/golden-test/documents/docx/ikujidetail_002197815.docx')
w = win32.DispatchEx('Word.Application'); w.Visible=False
try:
    d = w.Documents.Open(DOCX, ReadOnly=True)
    for target_i in (199, 5):  # i=199 over-wraps; i=5 is a normal body para for baseline
        p = d.Paragraphs(target_i)
        rng = p.Range
        txt = rng.Text
        n = min(rng.Characters.Count, 46)
        xs = []
        for k in range(1, n+1):
            c = rng.Characters(k)
            x = c.Information(5)  # horiz pos rel to page (pt)
            xs.append((c.Text, x))
        print('=== para i=%d, chars=%d, text=%r ==='%(target_i, rng.Characters.Count, txt[:50]))
        prev=None
        for ch,x in xs:
            adv = (x-prev) if prev is not None else 0
            disp = ch.replace('\r','\r').replace('\x0b','\v').replace('\t','\t')
            print('  %-3s x=%7.2f adv=%6.2f'%(repr(disp)[1:-1][:2], x, adv))
            prev=x
    d.Close(False)
finally:
    w.Quit()
