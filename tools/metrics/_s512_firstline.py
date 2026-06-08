# -*- coding: utf-8 -*-
"""S512 first-line offset mechanism: for each doc, the first few content lines' fs + Word
baseline + Oxi baseline + offset. See if the Oxi-high first-line offset scales with fs
(ascent-proportional) or is fixed. cp932-safe. Usage: python _s512_firstline.py <tag> <w.json> <oxi.json>"""
import json, io, sys


def lines(p, k):
    g = json.load(io.open(p, encoding='utf-8'))['pages'][0]['glyphs']
    rows = {}
    for x in g:
        if x['char'].strip():
            rows.setdefault(round(x[k], 1), []).append(x)
    ys = sorted(rows); out = []
    last = -99
    for y in ys:
        if y - last < 4:
            continue
        last = y
        r = sorted(rows[y], key=lambda c: c['x'])
        fs = max(c.get('fs', c.get('font_size', 0)) for c in r)
        out.append((y, fs))
    return out


tag = sys.argv[1]
W = lines(sys.argv[2], 'y')
O = lines(sys.argv[3], 'baseline')
print('=== %s: first 5 content lines ===' % tag)
print('  idx  Word_y  Word_fs | Oxi_y  Oxi_fs | offset(Oxi-Word)')
for i in range(min(5, len(W), len(O))):
    wy, wfs = W[i]; oy, ofs = O[i]
    print('  %d   %7.2f %5.1f | %7.2f %5.1f |  %+.2f' % (i, wy, wfs, oy, ofs, oy - wy))
