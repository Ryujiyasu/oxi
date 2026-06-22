# -*- coding: utf-8 -*-
import json, sys
sys.stdout.reconfigure(encoding='utf-8',errors='replace')
W=json.load(open('../../pipeline_data/pagination_word/tokyoshugyo.json',encoding='utf-8'))
OX=json.load(open(r'C:/tmp/tks_dump.json',encoding='utf-8'))

# Oxi ordered para list: (oxi_page, y, text)
opar=[]
for p in OX['pages']:
    pg=p['page']; groups={}; order=[]
    for e in p['elements']:
        if e.get('type') not in ('text','border') or not e.get('text'): continue
        key=(e.get('para_idx'),e.get('cell_para_idx'),e.get('cell_row_idx'),e.get('cell_col_idx'),round(e['y'],0))
        if key not in groups: groups[key]=[e['y'],'']; order.append(key)
        groups[key][1]+=e['text']
    for k in order: y,t=groups[k]; opar.append((pg,y,t))

# Word paras with text, sorted by i
wpar=[(w['i'],w['page'],w['y'],w['text']) for w in W['paragraphs'] if (w['text'] or '').strip()]

# match in order
oi=0; recs=[]
for wi,wpg,wy,wt in wpar:
    pref=wt[:8]; j=oi; found=None
    while j<len(opar) and j<oi+80:
        if opar[j][2][:8]==pref or (len(pref)>=6 and pref in opar[j][2][:24]):
            found=j; break
        j+=1
    if found is not None:
        opg,oyy,_=opar[found]; recs.append((wi,wpg,wy,opg,oyy)); oi=found+1

# For each WORD page in chapter, first matched para -> oxi page & y; report pdelta & y-drift
print(' Wpg | first-para oxiPg pdelta  Wy->Oy  ydrift')
seen=set()
for wi,wpg,wy,opg,oyy in recs:
    if wpg in seen or not(44<=wpg<=66): continue
    seen.add(wpg)
    print(f' {wpg:>3} | i{wi:<5} oxiPg{opg:>3}  d={opg-wpg:+d}  {wy:6.1f}->{oyy:6.1f}  yd={oyy-wy:+6.1f}')
