# -*- coding: utf-8 -*-
"""roudoujoken form-table per-row rendered height, Word vs Oxi. Word: pagination
JSON cell_row/page/y. Oxi: dump cell_row_idx/page/y. Compute each row's TOP
(page,y) on each side; within-page consecutive diffs = row heights. Cumulative
page offset to compare. Find the row(s) Oxi under-counts."""
import json,sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
USABLE=841.9-22.7-11.65  # roudoujoken content height ~807.55
TOP=22.7
W=json.load(open(r'pipeline_data/pagination_word/roudoujoken.json',encoding='utf-8'))
d=json.load(open(r'C:/Users/ryuji/AppData/Local/Temp/roudou_layout.json',encoding='utf-8'))
# Word: cell_row -> (page, min_y). Only the FIRST table (table_start of pages1-2).
wrow={}
for p in W['paragraphs']:
    if not p['in_table'] or p['cell_row'] is None: continue
    if p['page']>2: continue
    k=p['cell_row']
    if k not in wrow or (p['page'],p['y'])<(wrow[k][0],wrow[k][1]):
        wrow[k]=(p['page'],p['y'])
# Oxi: cell_row_idx -> (page, min_y)
orow={}
for pgno,pg in enumerate(d['pages']):
    if pgno>1: continue
    for e in pg['elements']:
        if e['type']!='text' or not e['text'].strip(): continue
        k=e.get('cell_row_idx')
        if k is None: continue
        if k not in orow or (pgno+1,e['y'])<(orow[k][0],orow[k][1]):
            orow[k]=(pgno+1,e['y'])
def absy(pg,y): return (pg-1)*USABLE + (y-TOP)
print('row | Word(pg,y -> abs)      | Oxi(pg,y -> abs)       | absΔ  rowH_W rowH_O dH')
keys=sorted(set(wrow)|set(orow))
pW=pO=None
for k in keys:
    w=wrow.get(k); o=orow.get(k)
    wa=absy(*w) if w else None; oa=absy(*o) if o else None
    rhW=(wa-pW) if (wa is not None and pW is not None) else None
    rhO=(oa-pO) if (oa is not None and pO is not None) else None
    dH=(rhO-rhW) if (rhW is not None and rhO is not None) else None
    print('r%-2d | %-22s | %-22s | %s %s %s %s'%(
        k,
        ('p%d y%6.1f a%6.1f'%(w[0],w[1],wa)) if w else '   -   ',
        ('p%d y%6.1f a%6.1f'%(o[0],o[1],oa)) if o else '   -   ',
        ('%+5.1f'%(oa-wa)) if (wa is not None and oa is not None) else '  -  ',
        ('%5.1f'%rhW) if rhW is not None else '  -  ',
        ('%5.1f'%rhO) if rhO is not None else '  -  ',
        ('%+5.1f'%dH) if dH is not None else '  -  '))
    if wa is not None: pW=wa
    if oa is not None: pO=oa
