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
    ('repro_pcw_V12.docx','V11 + 1636-style docDefault rFonts'),
    ('repro_pcw_V13.docx','V11 + indent (leftChars=150 left=315 right=199)'),
    ('repro_pcw_V14.docx','V11 + para-level <w:wordWrap/> (overriding style off)'),
    ('repro_pcw_V15.docx','V11 + style widowControl=0 wordWrap=0 adjustRightInd=0'),
    ('repro_pcw_V16.docx','V11 + style explicit rFonts ascii/hAnsi/cs=Mincho'),
    ('repro_pcw_V17.docx','V11 + multi-run split (some w/o hint=eastAsia)'),
    ('repro_pcw_V18.docx','V11 + ALL real-1636 properties combined'),
    ('repro_pcw_V19.docx','V11 + settings:balanceSingleByteDoubleByteWidth'),
    ('repro_pcw_V20.docx','V11 + settings:useFELayout'),
    ('repro_pcw_V21.docx','V11 + settings:characterSpacingControl=compressPunctuation'),
    ('repro_pcw_V22.docx','V11 + ALL 3 settings flags'),
    ('repro_pcw_V23.docx','V18 + ALL settings flags (full mimicry)'),
    ('repro_pcw_V24.docx','BALANCE + cs=0 (isolate base)'),
    ('repro_pcw_V25.docx','BALANCE + cs=-20 (-1pt)'),
    ('repro_pcw_V26.docx','BALANCE + cs=+20 (+1pt)'),
    ('repro_pcw_V27.docx','BALANCE + cs=-5 (-0.25pt)'),
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
