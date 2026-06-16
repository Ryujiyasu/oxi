# -*- coding: utf-8 -*-
"""Per-char advances of a Word PDF line vs Oxi natural, for the で定める。 line."""
import os, sys, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz
PDF = os.path.join(tempfile.gettempdir(), 'tks_truth.pdf')
doc = fitz.open(PDF)
pg = doc[43]  # page 44
rd = pg.get_text('rawdict')
for blk in rd['blocks']:
    if blk.get('type',0)!=0: continue
    for ln in blk.get('lines',[]):
        chars=[]
        for sp in ln['spans']:
            for ch in sp['chars']:
                chars.append((ch['c'], ch['bbox']))
        txt=''.join(c for c,_ in chars)
        if '取扱いについては' not in txt: continue
        print(f"WORD line ({len(chars)} chars): x0={chars[0][1][0]:.1f}")
        prev=None
        for c,b in chars:
            adv = (b[0]-prev) if prev is not None else 0  # start-to-start
            mark=''
            if c in '、。，．・「」『』（）':
                mark=' <-yakumono'
            print(f"  '{c}' x[{b[0]:6.1f}..{b[2]:6.1f}] w={b[2]-b[0]:5.1f} adv={adv:5.1f}{mark}")
            prev=b[0]
