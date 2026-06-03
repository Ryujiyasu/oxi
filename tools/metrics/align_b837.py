# -*- coding: utf-8 -*-
"""S492c clean — Oxi-vs-Word para alignment for b837, authored as a UTF-8 FILE
(bash heredocs mangle Japanese literals on this cp932 system). Finds the Oxi para
containing 本府 / オープンデータに関する事業 and reports its line structure + page span,
to compare against the Word ground truth (本府は para = 4 lines [34,35,35,14], 2 on p6 / 2 on p7).
"""
import json

d = json.load(open('c:/tmp/_b837.json', encoding='utf-8'))
paras = {}
for pi, pg in enumerate(d['pages']):
    for e in pg['elements']:
        if e['type'] == 'text' and e.get('para_idx') is not None:
            paras.setdefault(e['para_idx'], []).append((pi, e))

needles = ['本府', 'オープンデータに関する事業', '内閣官房', '公開が望まれる分野']
allt = ''.join(e['text'] for pg in d['pages'] for e in pg['elements'] if e['type'] == 'text')
print("Oxi total chars:", len(allt))
for nd in needles:
    print("  %r in Oxi full text? %s" % (nd, nd in allt))

for nd in ['本府', 'オープンデータに関する事業']:
    print("\n=== Oxi paras containing %r ===" % nd)
    for paidx in sorted(paras):
        els = sorted(paras[paidx], key=lambda t: (t[0], round(t[1]['y'], 1), t[1]['x']))
        txt = ''.join(e['text'] for _, e in els)
        if nd in txt:
            from collections import OrderedDict
            lines = OrderedDict()
            for pi, e in els:
                lines.setdefault((pi, round(e['y'], 1)), []).append(e)
            counts = []
            pages = []
            for (pi, y), ln in lines.items():
                counts.append(sum(len(e['text']) for e in ln))
                pages.append(pi + 1)
            print("  para_idx %d len=%d start=%r" % (paidx, len(txt), txt[:16]))
            print("     lines=%d counts=%s pages=%s" % (len(lines), counts, pages))
            break
