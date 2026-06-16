# -*- coding: utf-8 -*-
"""Extract Word PDF char positions for a given page, lines containing a substring."""
import os, sys, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz
PDF = os.path.join(tempfile.gettempdir(), 'tks_truth.pdf')
doc = fitz.open(PDF)
page = int(sys.argv[1]); needle = sys.argv[2] if len(sys.argv) > 2 else None
pg = doc[page-1]
rd = pg.get_text('rawdict')
for blk in rd['blocks']:
    if blk.get('type',0)!=0: continue
    for ln in blk.get('lines',[]):
        chars=[]
        for sp in ln['spans']:
            for ch in sp['chars']:
                chars.append((ch['c'], ch['bbox']))
        txt=''.join(c for c,_ in chars)
        if needle and needle not in txt: continue
        if not txt.strip(): continue
        x0=min(b[0] for _,b in chars); x1=max(b[2] for _,b in chars)
        y0=min(b[1] for _,b in chars)
        print(f"y{y0:6.1f} x[{x0:6.1f}..{x1:6.1f}] | {txt}")
        # last 4 chars positions
        for c,b in chars[-4:]:
            print(f"      '{c}' x[{b[0]:.1f}..{b[2]:.1f}] w={b[2]-b[0]:.1f}")
