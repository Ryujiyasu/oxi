# -*- coding: utf-8 -*-
import json, sys
sys.stdout.reconfigure(encoding='utf-8',errors='replace')
PH=841.92
W=json.load(open('../../pipeline_data/pagination_word/tokyoshugyo.json',encoding='utf-8'))
OX=json.load(open(r'C:/tmp/tks_dump.json',encoding='utf-8'))

# Oxi: build ordered para list (abs_y, text) by para_idx within page, body+cell
opar=[]
for p in OX['pages']:
    pg=p['page']
    groups={}
    order=[]
    for e in p['elements']:
        if e.get('type') not in ('text','border'): continue
        if not e.get('text'): continue
        key=(e.get('para_idx'), e.get('cell_para_idx'), e.get('cell_row_idx'), e.get('cell_col_idx'), round(e['y'],0))
        if key not in groups:
            groups[key]=[e['y'], '']; order.append(key)
        groups[key][1]+=e['text']
    for k in order:
        y,t=groups[k]
        opar.append(((pg-1)*PH+y, t))

# Word: chapter paras with text
wpar=[(w['i'],(w['page']-1)*PH+w['y'], w['text'], w['page']) for w in W['paragraphs'] if (w['text'] or '').strip()]

# match each word para to next oxi para (in order) with same 8-char prefix
oi=0; rows=[]
for wi,wy,wt,wpg in wpar:
    pref=wt[:8]
    j=oi; found=None
    while j<len(opar) and j<oi+60:
        if opar[j][1][:8]==pref or (len(pref)>=6 and pref in opar[j][1][:20]):
            found=j; break
        j+=1
    if found is not None:
        oy=opar[found][0]
        rows.append((wi,wpg,wy,oy,oy-wy,wt[:18]))
        oi=found+1

# print offset every ~15 chapter paras (i in 1439..1980)
print('  wi  wpg   wordY    oxiY   offset  text')
prev=None
for wi,wpg,wy,oy,off,t in rows:
    if 1430<=wi<=1990:
        if prev is None or abs(off-prev)>=6 or wi%40==0:
            print(f'{wi:>5} {wpg:>4} {wy:8.1f} {oy:8.1f} {off:+8.1f}  {t}')
            prev=off
