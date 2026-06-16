# -*- coding: utf-8 -*-
"""Per-para LINE COUNT Word(COM Y-gap) vs Oxi(raw dump distinct Y rows), 賃金
chapter. Flags paras where Oxi has FEWER lines than Word (= the over-fit source)."""
import json,sys
sys.stdout.reconfigure(encoding='utf-8',errors='replace')
W=json.load(open('pipeline_data/pagination_word/tokyoshugyo.json',encoding='utf-8'))
d=json.load(open('C:/tmp/tks_base.json',encoding='utf-8'))
def norm(t): return (t or '').replace('　','').replace(' ','').strip()
# Oxi: group raw text elements by (para_idx, cell ids) -> distinct Y rows + text
from collections import defaultdict
groups=defaultdict(lambda: {'ys':set(),'txt':[]})
for pg in d['pages']:
    for el in pg.get('elements',[]):
        if el.get('type')!='text': continue
        key=(el.get('para_idx'),el.get('cell_para_idx'),el.get('cell_row_idx'),el.get('cell_col_idx'),round(el['y']/200))
        groups[key]['ys'].add(round(el['y'],0))
        groups[key]['txt'].append((el['y'],el['x'],el.get('text','')))
oxi={}
for g in groups.values():
    g['txt'].sort()
    t=norm(''.join(x[2] for x in g['txt']))
    if t and t not in oxi: oxi[t]=len(g['ys'])
# Word paras with text in 賃金 chapter, line count = gap/18 (or 12 for dense)
paras={r['i']:r for r in W['paragraphs']}
flagged=[]
for i in range(1439,1965):
    r=paras.get(i)
    if not r or not norm(r['text']): continue
    nxt=next((paras[j] for j in range(i+1,i+5) if j in paras), None)
    if not nxt or nxt['page']!=r['page']: continue
    gap=nxt['y']-r['y']
    if not (0<gap<200): continue
    wlines=max(1,round(gap/18.0))
    k=norm(r['text'])[:12]
    om=[v for t,v in oxi.items() if t.startswith(k)]
    if om:
        ol=om[0]
        if ol < wlines:
            flagged.append((i,wlines,ol,r['text'][:34]))
print(f"paras where Oxi has FEWER lines than Word (over-fit), 賃金 chapter:")
tot=0
for i,wl,ol,t in flagged:
    print(f"  i{i} W{wl}/O{ol} (-{wl-ol}) | {t}")
    tot+=wl-ol
print(f"TOTAL Oxi line deficit (flagged): {tot}")
