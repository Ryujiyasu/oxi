# -*- coding: utf-8 -*-
"""Scan a Word PDF page's lines: per line, report mid-line 約物 advances (detect
DEMAND oikomi = 約物 compressed below the unconditional ×0.5 kern). Tests whether
Word oikomi is regional (region-2 only) or global."""
import os, sys, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz
PDF = os.path.join(tempfile.gettempdir(), 'tks_truth.pdf')
doc = fitz.open(PDF)
YAK = set('、。，．')  # commas/periods (the demand-compressed class)
for pageno in [int(a) for a in sys.argv[1:]]:
    pg = doc[pageno-1]
    rd = pg.get_text('rawdict')
    n_lines = n_oikomi = 0
    examples = []
    for blk in rd['blocks']:
        if blk.get('type',0)!=0: continue
        for ln in blk.get('lines',[]):
            chars=[]
            for sp in ln['spans']:
                for ch in sp['chars']:
                    chars.append((ch['c'], ch['bbox']))
            txt=''.join(c for c,_ in chars).strip()
            if len(txt)<10: continue
            n_lines += 1
            # mid-line comma/period advances (exclude last char)
            comp=[]
            for i in range(len(chars)-1):
                c,b=chars[i]; _,nb=chars[i+1]
                if c in YAK:
                    adv=nb[0]-b[0]
                    comp.append((c,round(adv,1)))
            # demand oikomi = a 、。 with advance < 8.0 (below ×0.5 kern ~5.25..half; natural ~10.5)
            tight=[a for _,a in comp if a < 8.0]
            if tight:
                n_oikomi += 1
                if len(examples)<4: examples.append((txt[:36], comp))
    print(f"p{pageno}: {n_lines} lines, {n_oikomi} with a DEMAND-compressed mid 、。 ({100*n_oikomi//max(1,n_lines)}%)")
    for t,c in examples: print(f"    {t}  {c}")
