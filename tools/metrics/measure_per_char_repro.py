# -*- coding: utf-8 -*-
"""COM-measure per-char advance for the per_char_repro variants.

For each variant, measures Word's per-character HPOS via
Range(Start+i, Start+i).Information(5) and computes mean advance.
"""
import os, sys, json
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

variants = [
    ('repro_pcw_V1.docx', 'cs=0       no kern'),
    ('repro_pcw_V2.docx', 'cs=-9      no kern'),
    ('repro_pcw_V3.docx', 'cs=0       kern=2'),
    ('repro_pcw_V4.docx', 'cs=-9      kern=2'),
    ('repro_pcw_V5.docx', 'cs=-9 k=2  noASDE/DN'),
    ('repro_pcw_V6.docx', 'cs=-9 k=2  noASDE/DN snapGrid=0'),
    ('repro_pcw_V7.docx', 'cs=-9 k=2  noASDE/DN snapGrid=0 jc=both'),
    ('repro_pcw_V8.docx', 'cs=0 no docGrid (baseline)'),
    ('repro_pcw_V9.docx', 'TABLE cs=-9 snapGrid=0 (1636-like)'),
    ('repro_pcw_V10.docx','TABLE cs=-9 snapGrid=0 + style.cs=-1'),
    ('repro_pcw_V11.docx','TABLE cs=-9 snapGrid=0 + style.cs=-1 + jc=both'),
]

word = wc.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
results = []
for name, label in variants:
    path = os.path.abspath(f'tools/metrics/_repros/{name}')
    if not os.path.exists(path):
        print(f'  SKIP (missing): {name}')
        continue
    print(f'\n=== {name}  {label} ===')
    d = word.Documents.Open(path, ReadOnly=True)
    try:
        p = d.Paragraphs(1)
        rng = p.Range
        text = rng.Text or ''
        chars = []
        for off in range(min(len(text), 30)):
            ch = text[off]
            if ch in ('\r', '\n', '\x07'):
                continue
            cr = d.Range(rng.Start + off, rng.Start + off)
            x = cr.Information(5)
            y = cr.Information(6)
            chars.append({'off': off, 'ch': ch, 'x': x, 'y': y})
        # advances on same line
        advances = []
        for j in range(1, len(chars)):
            if chars[j]['y'] == chars[j-1]['y']:
                advances.append(chars[j]['x'] - chars[j-1]['x'])
        if advances:
            mean = sum(advances) / len(advances)
            print(f'  text_chars={len(chars)}  advances={[round(a,3) for a in advances[:12]]}')
            print(f'  mean={mean:.4f}pt   n={len(advances)}')
            results.append({'variant': name, 'label': label, 'mean': mean, 'advances': advances, 'first_x': chars[0]['x']})
    finally:
        d.Close(SaveChanges=False)
word.Quit()

json.dump(results, open('pipeline_data/per_char_repro_results.json', 'w', encoding='utf-8'),
          ensure_ascii=False, indent=2)
print('\nSaved pipeline_data/per_char_repro_results.json')
