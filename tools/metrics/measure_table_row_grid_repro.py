# -*- coding: utf-8 -*-
"""COM-measure row Y positions in table-row-height repro variants."""
import os, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

variants = [
    ('repro_trh_RT1.docx', 'type=lines pitch=330 line=300exact (bd90b00 mimic)'),
    ('repro_trh_RT2.docx', 'type=linesAndChars pitch=330 line=300exact'),
    ('repro_trh_RT3.docx', 'type=lines pitch=330 line=240exact (small)'),
    ('repro_trh_RT4.docx', 'type=lines pitch=330 line=400exact (large)'),
    ('repro_trh_RT5.docx', 'type=lines pitch=330 line=auto/single (bd90b00 cell mimic)'),
    ('repro_trh_RT6.docx', 'type=linesAndChars pitch=330 line=auto/single'),
    ('repro_trh_RT7.docx', 'type=lines pitch=240 line=auto/single (12pt grid)'),
]

word = wc.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
for name, label in variants:
    path = os.path.abspath(f'tools/metrics/_repros/{name}')
    if not os.path.exists(path):
        print(f'  SKIP: {name}')
        continue
    print(f'\n=== {name}  {label} ===')
    d = word.Documents.Open(path, ReadOnly=True)
    try:
        ys = []
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            rng = p.Range
            text = (rng.Text or '').replace('\r', '\\r').replace('\x07', '')
            if '行' not in text:
                continue
            cr = d.Range(rng.Start, rng.Start)
            y = cr.Information(6)
            ys.append((i, y, text[:20]))
        print(f'  Row Y positions:')
        prev = None
        for i, y, t in ys:
            gap = '' if prev is None else f' Δ={y - prev:.2f}'
            print(f'    i={i} y={y:.2f}  text={t!r:25s} {gap}')
            prev = y
    finally:
        d.Close(SaveChanges=False)
word.Quit()
