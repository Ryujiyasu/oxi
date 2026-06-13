# -*- coding: utf-8 -*-
"""S559 wheel-(a) lever — localize kyotei36spec over-pagination (Oxi 5 pages vs
Word 4, delta_hist {-1:35}). Walk Word paragraphs in document order, match to Oxi
by normalized text, show Word-page vs Oxi-page and flag where the delta first
appears + the table/row context, so the over-height block can be found.
"""
import json
import sys
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')
W = json.load(open(r'pipeline_data/pagination_word/kyotei36spec.json', encoding='utf-8'))
O = json.load(open(r'pipeline_data/pagination_oxi/kyotei36spec.json', encoding='utf-8'))


def norm(s):
    return ''.join((s or '').split())[:16]


# oxi text -> list of (page) in doc order
oxi_recs = []
for pg in sorted(O['pages'], key=lambda x: int(x)):
    for r in O['pages'][pg]:
        oxi_recs.append((int(pg), norm(r.get('text', '')), r))
oxi_by_t = defaultdict(list)
for pgi, (page, t, r) in enumerate(oxi_recs):
    oxi_by_t[t].append((page, r))

used = defaultdict(int)
rows = []
for p in W['paragraphs']:
    wt = norm(p.get('text', ''))
    wp = p['page']
    cand = oxi_by_t.get(wt)
    op = None
    rec = None
    if cand and wt:
        idx = used[wt]
        if idx < len(cand):
            op, rec = cand[idx]
            used[wt] += 1
        else:
            op, rec = cand[-1]
    rows.append((p['i'], wp, op, p.get('in_table'), p.get('cell_row'), p.get('cell_col'), p.get('table_start'), wt))

print('Word pages=%d  Oxi pages=%d' % (W['n_pages'], O['n_pages']))
print('\ni    Wpg Opg d  tbl r,c tstart  text')
prev_d = 0
for i, wp, op, intbl, cr, cc, ts, t in rows:
    d = (op - wp) if op is not None else None
    mark = ''
    if d is not None and d != prev_d:
        mark = '  <<< delta changes %s->%s' % (prev_d, d)
        prev_d = d
    ds = ('%+d' % d) if d is not None else '?'
    # print only the transition region and table starts
    if mark or (d not in (0, None) and i % 1 == 0):
        print('%-4d %-3s %-3s %-2s %s %s,%s %s  %s%s'
              % (i, wp, op if op is not None else '-', ds, 'T' if intbl else '.',
                 cr, cc, ts, t, mark))
