# -*- coding: utf-8 -*-
"""S492c final — is the 本府は paragraph in a TABLE in Oxi? Search ALL text elements
(including table cells, which have para_idx=None / cell_* set). If 本府は is a cell,
the b837 p7 +18 is a TABLE/structural issue (S451/S452 closed frontier), NOT body
line-breaking — and my whole char-width chase was on the wrong (body) paragraph.
"""
import json

d = json.load(open('c:/tmp/_b837.json', encoding='utf-8'))
hits = []
for pi, pg in enumerate(d['pages']):
    for e in pg['elements']:
        if e['type'] != 'text':
            continue
        if '本府' in e['text'] or 'オープンデータに関する' in e['text'] or '内閣官房' in e['text']:
            hits.append((pi + 1, e))
print("elements containing 本府/オープンデータに関する/内閣官房:", len(hits))
for pgn, e in hits[:12]:
    print("  page%d y=%.1f x=%.1f para_idx=%s cell_para_idx=%s cell_row=%s cell_col=%s text=%r"
          % (pgn, e['y'], e['x'], e.get('para_idx'), e.get('cell_para_idx'),
             e.get('cell_row_idx'), e.get('cell_col_idx'), e['text'][:24]))

# Also: count how many text elements are cell (para_idx None) vs body
ncell = sum(1 for pg in d['pages'] for e in pg['elements']
            if e['type'] == 'text' and e.get('para_idx') is None)
nbody = sum(1 for pg in d['pages'] for e in pg['elements']
            if e['type'] == 'text' and e.get('para_idx') is not None)
print("\nOxi b837 text elements: body(para_idx set)=%d, cell/other(para_idx None)=%d" % (nbody, ncell))
