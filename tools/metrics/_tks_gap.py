# -*- coding: utf-8 -*-
"""Compare consecutive para Y-gaps Word(COM) vs Oxi(dump) on a given page.
Matches paras by text prefix within the page. Reliable (no fitz)."""
import json, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
W = json.load(open('pipeline_data/pagination_word/tokyoshugyo.json', encoding='utf-8'))
O = json.load(open('pipeline_data/pagination_oxi/tokyoshugyo.json', encoding='utf-8'))
wp = {}
for r in W['paragraphs']:
    wp.setdefault(r['page'], []).append(r)
op = O['pages']

def keyt(t):
    return (t or '').strip()[:8]

page = int(sys.argv[1])
wlist = sorted(wp.get(page, []), key=lambda r: r['y'])
olist = sorted(op.get(str(page), []), key=lambda r: (r['y'], r['x']))
print(f"=== page {page}: Word {len(wlist)} paras, Oxi {len(olist)} recs ===")
# build oxi text->y lookup
from collections import defaultdict
olook = defaultdict(list)
for r in olist:
    olook[keyt(r['text'])].append(r['y'])
print(f"{'Wy':>7} {'Oy':>7} | {'Wgap':>6} {'Ogap':>6} {'dGap':>6} | text")
prevw = prevo = None
for r in wlist:
    k = keyt(r['text'])
    oy = None
    if k and k in olook and olook[k]:
        oy = olook[k].pop(0)
    wy = r['y']
    wg = (wy - prevw) if prevw is not None else 0
    og = (oy - prevo) if (oy is not None and prevo is not None) else 0
    dg = (og - wg) if (wg and og) else 0
    flag = ''
    if dg < -2: flag = ' <<< Oxi gap SHORT'
    elif dg > 2: flag = ' >>> Oxi gap LONG'
    txt = (r['text'] or '')[:30]
    print(f"{wy:>7.1f} {(oy if oy is not None else 0):>7.1f} | {wg:>6.1f} {og:>6.1f} {dg:>+6.1f} |{flag} {txt}")
    prevw = wy
    if oy is not None:
        prevo = oy
