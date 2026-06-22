# -*- coding: utf-8 -*-
import json, sys
sys.stdout.reconfigure(encoding='utf-8',errors='replace')
W=json.load(open('../../pipeline_data/pagination_word/tokyoshugyo.json',encoding='utf-8'))
OX=json.load(open(r'C:/tmp/tks_dump.json',encoding='utf-8'))
LO=int(sys.argv[1]); HI=int(sys.argv[2])
opar=[]
for p in OX['pages']:
    pg=p['page']; groups={}; order=[]
    for e in p['elements']:
        if e.get('type') not in ('text','border') or not e.get('text'): continue
        key=(e.get('para_idx'),e.get('cell_para_idx'),e.get('cell_row_idx'),e.get('cell_col_idx'),round(e['y'],0))
        if key not in groups: groups[key]=[e['y'],'']; order.append(key)
        groups[key][1]+=e['text']
    for k in order: y,t=groups[k]; opar.append((pg,y,t))
wpar=[(w['i'],w['page'],w['y'],w['text']) for w in W['paragraphs'] if (w['text'] or '').strip()]
oi=0
for wi,wpg,wy,wt in wpar:
    pref=wt[:8]; j=oi; found=None
    while j<len(opar) and j<oi+80:
        if opar[j][2][:8]==pref or (len(pref)>=6 and pref in opar[j][2][:24]): found=j; break
        j+=1
    if found is not None:
        opg,oyy,_=opar[found]; oi=found+1
        if LO<=wi<=HI:
            d=(opg-wpg)*841.92+(oyy-wy)
            print(f'{wi:>5} W{wpg:>3} {wy:6.1f} | O{opg:>3} {oyy:6.1f}  {d:+7.1f} | {wt[:32]}')
