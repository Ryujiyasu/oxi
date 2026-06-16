# -*- coding: utf-8 -*-
"""Region-2 per-para Y-gap (height) Word(COM) vs Oxi(dump). For consecutive
matched paras ON THE SAME Word page, compare gap = Y[i+1]-Y[i]. Sums per page.
Where Word gap > Oxi gap, Oxi under-counts height (spacing/row/line)."""
import json, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
W = json.load(open('pipeline_data/pagination_word/tokyoshugyo.json', encoding='utf-8'))
O = json.load(open('pipeline_data/pagination_oxi/tokyoshugyo.json', encoding='utf-8'))
def norm(t): return (t or '').replace('　','').replace(' ','').strip()
oxi = {}
for pg, recs in O['pages'].items():
    for r in recs:
        k = norm(r['text'])
        if k and k not in oxi: oxi[k] = (int(pg), r['y'])
wp = W['paragraphs']
gap_short = 0.0
print(f"{'wpg':>3} {'Wsum':>7} {'Osum':>7} {'diff':>7} {'npairs':>6}  (Wsum-Osum; +=Oxi gaps SHORTER=under-count)")
for pgnum in range(46, 64):
    paras = sorted([r for r in wp if r['page']==pgnum and norm(r['text'])], key=lambda r:r['y'])
    wsum = osum = 0.0; n = 0
    for a, b in zip(paras, paras[1:]):
        ka, kb = norm(a['text']), norm(b['text'])
        if ka in oxi and kb in oxi and oxi[ka][0]==oxi[kb][0]:  # same Oxi page
            wg = b['y'] - a['y']; og = oxi[kb][1] - oxi[ka][1]
            if 0 < wg < 200 and 0 < og < 200:
                wsum += wg; osum += og; n += 1
    diff = wsum - osum; gap_short += diff
    print(f"{pgnum:>3} {wsum:>7.1f} {osum:>7.1f} {diff:>+7.1f} {n:>6}")
print(f"TOTAL Wsum-Osum = {gap_short:+.1f}pt  (positive => Oxi systematically packs paras tighter)")
