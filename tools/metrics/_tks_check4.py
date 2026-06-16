# -*- coding: utf-8 -*-
"""For each of the 4 S586-fired paras, find its lines in the Word PDF and show
where Word breaks (does Word fit the char or wrap it = oikomi vs oidashi)."""
import os, sys, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz
PDF = os.path.join(tempfile.gettempdir(), 'tks_truth.pdf')
doc = fitz.open(PDF)
needles = ['休職期間中に休職事由が消滅', '育児・介護休業等の取扱い',
           '労働者の定年は、満', '次のいずれかに該当するときは']
def lines_with(needle):
    out=[]
    for pi in range(len(doc)):
        for blk in doc[pi].get_text('rawdict')['blocks']:
            if blk.get('type',0)!=0: continue
            for ln in blk.get('lines',[]):
                t=''.join(c['c'] for sp in ln['spans'] for c in sp['chars'])
                if needle in t:
                    # show this line + the NEXT line (to see the break)
                    out.append((pi+1, t))
    return out
for nd in needles:
    print(f"=== {nd} ===")
    res=lines_with(nd)
    for pg,t in res[:2]:
        print(f"  p{pg}: {t}")
