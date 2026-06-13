# -*- coding: utf-8 -*-
"""Localize roudoujoken's single -1 delta (FAIL 0.993)."""
import json,sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
W=json.load(open(r'pipeline_data/pagination_word/roudoujoken.json',encoding='utf-8'))
O=json.load(open(r'pipeline_data/pagination_oxi/roudoujoken.json',encoding='utf-8'))
def norm(s): return ''.join((s or '').split())[:16]
oxi_by_t=defaultdict(list)
for pg in sorted(O['pages'],key=lambda x:int(x)):
    for r in O['pages'][pg]: oxi_by_t[norm(r.get('text',''))].append(int(pg))
used=defaultdict(int); prev_d=0
print('Word pages=%d Oxi pages=%d'%(W['n_pages'],O['n_pages']))
print('i    Wpg Opg d   tbl r,c  text')
for p in W['paragraphs']:
    wt=norm(p.get('text','')); wp=p['page']
    cand=oxi_by_t.get(wt); op=None
    if cand and wt:
        idx=used[wt]; op=cand[idx] if idx<len(cand) else cand[-1]; used[wt]+=1
    d=(op-wp) if op is not None else None
    mark=''
    if d is not None and d!=prev_d:
        mark='  <<< %s->%s'%(prev_d,d); prev_d=d
    if mark or (d not in (0,None)):
        print('%-4d %-3s %-3s %-3s %s %s,%s  %s%s'%(p['i'],wp,op if op is not None else '-',
              ('%+d'%d) if d is not None else '?','T' if p['in_table'] else '.',
              p['cell_row'],p['cell_col'],wt,mark))
