# -*- coding: utf-8 -*-
"""Scan Word PDF: every 、 (or 。) immediately followed by an OPENING bracket —
report its advance. Tests if 、「 collapse to ~3.0 is UNCONDITIONAL (always) or
page-44-specific (demand)."""
import os, sys, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz
PDF = os.path.join(tempfile.gettempdir(), 'tks_truth.pdf')
doc = fitz.open(PDF)
OPEN = set('「『（〔【《〈｛［')
advs=[]
for pageno in range(1, len(doc)+1):
    pg = doc[pageno-1]
    for blk in pg.get_text('rawdict')['blocks']:
        if blk.get('type',0)!=0: continue
        for ln in blk.get('lines',[]):
            chars=[]
            for sp in ln['spans']:
                for ch in sp['chars']: chars.append((ch['c'], ch['bbox']))
            for i in range(len(chars)-1):
                c,b=chars[i]; nc,nb=chars[i+1]
                if c in '、。' and nc in OPEN:
                    adv=round(nb[0]-b[0],1)
                    advs.append((pageno, c, nc, adv))
import statistics
print(f"total 、/。 -> opening-bracket adjacencies: {len(advs)}")
vals=[a for _,_,_,a in advs]
if vals:
    print(f"  advance: min={min(vals)} max={max(vals)} mean={statistics.mean(vals):.1f} median={statistics.median(vals)}")
    # histogram buckets
    from collections import Counter
    buck=Counter(round(a) for a in vals)
    print("  advance histogram (rounded):", dict(sorted(buck.items())))
    print("  samples:", advs[:12])
