# -*- coding: utf-8 -*-
"""Inspect 2ea81a paragraphs around the page break (i=118 → i=119) via COM."""
import os, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

path = os.path.abspath('tools/golden-test/documents/docx/2ea81a8441cc_0025006-192.docx')
word = wc.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
d = word.Documents.Open(path, ReadOnly=True)
try:
    paras = d.Paragraphs
    n = paras.Count
    print(f'Total paragraphs: {n}')
    for i in range(115, min(n+1, 126)):
        p = paras(i)
        rng = p.Range
        start = d.Range(rng.Start, rng.Start)
        end = d.Range(rng.End - 1, rng.End - 1) if rng.End > rng.Start else start
        page_s = start.Information(3)
        page_e = end.Information(3)
        y_s = start.Information(6)
        y_e = end.Information(6)
        text = (rng.Text or '').replace('\r', '\\r').replace('\x07', '\\x07')[:60]
        in_table = bool(p.Range.Tables.Count > 0)
        line_h = None
        try:
            ls = p.Format.LineSpacing
            line_h = float(ls)
        except Exception:
            pass
        try:
            sb = p.Format.SpaceBefore
            sa = p.Format.SpaceAfter
        except Exception:
            sb = sa = None
        try:
            pbb = p.Format.PageBreakBefore
        except Exception:
            pbb = None
        try:
            kn = p.Format.KeepWithNext
            kt = p.Format.KeepTogether
        except Exception:
            kn = kt = None
        print(f'i={i}  page_s={page_s} page_e={page_e}  y_s={y_s:.2f} y_e={y_e:.2f}  '
              f'in_table={in_table}  line={line_h}  sb={sb} sa={sa}  pbb={pbb} kn={kn} kt={kt}  text={text!r}')
finally:
    d.Close(SaveChanges=False)
word.Quit()
