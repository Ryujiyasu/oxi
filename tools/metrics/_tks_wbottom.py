# -*- coding: utf-8 -*-
"""Word PDF: lowest content-line TOP-Y per page (excl footer y>770), 賃金 chapter.
Compares to Oxi's ~732-738. If Word's max content-Y is consistently LOWER, Word's
effective content-bottom is higher -> Word breaks ~1 line earlier -> more pages."""
import os,sys,tempfile
sys.stdout.reconfigure(encoding='utf-8',errors='replace')
import fitz
doc=fitz.open(os.path.join(tempfile.gettempdir(),'tks_truth.pdf'))
print(f"{'Wpg':>3} {'maxY':>7} {'lastline':>30}")
for pg in range(46,65):
    p=doc[pg-1]; best=0; txt=''
    for blk in p.get_text('dict')['blocks']:
        if blk.get('type',0)!=0: continue
        for ln in blk.get('lines',[]):
            t=''.join(s['text'] for s in ln['spans']).strip()
            if not t: continue
            y0=min(s['bbox'][1] for s in ln['spans'])
            if y0>770: continue  # footer
            if y0>best: best=y0; txt=t
    print(f"{pg:>3} {best:>7.1f} {txt[:28]:>30}")
