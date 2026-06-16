# -*- coding: utf-8 -*-
"""Region-2 per-Word-page content-height comparison: Word COM para Y-gaps vs Oxi
dump para Y-gaps (matched by normalized text). Sum gaps per Word page; where
Word's sum > Oxi's, Oxi under-counts that page's height. Offset-independent."""
import json, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
W = json.load(open('pipeline_data/pagination_word/tokyoshugyo.json', encoding='utf-8'))
O = json.load(open('pipeline_data/pagination_oxi/tokyoshugyo.json', encoding='utf-8'))
def norm(t): return (t or '').replace('　','').replace(' ','').strip()
# Oxi: text(normalized) -> list of (page, y) in doc order
from collections import defaultdict
oxi = defaultdict(list)
for pg, recs in O['pages'].items():
    for r in recs:
        k = norm(r['text'])
        if k: oxi[k].append((int(pg), r['y']))
wp = W['paragraphs']
# per Word page: count paras, and Oxi page of matched paras
from collections import Counter
print(f"{'wpg':>3} {'Wpar':>5} {'matched':>7} {'OxiPgs':>20}")
for pgnum in range(46, 65):
    paras = [r for r in wp if r['page']==pgnum and norm(r['text'])]
    opages = []
    for r in paras:
        k = norm(r['text'])
        if k in oxi and oxi[k]:
            opages.append(oxi[k][0][0])
    c = Counter(opages)
    span = f"{min(opages)}-{max(opages)}" if opages else "-"
    print(f"{pgnum:>3} {len(paras):>5} {len(opages):>7}   Oxi pages {dict(sorted(c.items()))}")
