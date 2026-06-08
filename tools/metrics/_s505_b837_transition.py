# -*- coding: utf-8 -*-
"""S505 b837 p4 body->footnote transition: list Word vs Oxi lines (clustered by baseline y)
in y 520-670 to find the line/gap Word has that Oxi lacks (the −14pt 1-line cascade at
y631). cp932-safe: ASCII out file."""
import json, io
W = json.load(io.open('c:/tmp/b837_w.json', encoding='utf-8'))['pages'][3]['glyphs']
O = json.load(io.open('c:/tmp/b837_ox.json', encoding='utf-8'))['pages'][3]['glyphs']


def lines(glyphs, ykey, ylo, yhi):
    rows = {}
    for g in glyphs:
        if not g['char'].strip():
            continue
        y = g[ykey]
        if y < ylo or y > yhi:
            continue
        rows.setdefault(round(y, 0), []).append(g)
    out = []
    for k in sorted(rows):
        r = sorted(rows[k], key=lambda g: g['x'])
        fs = max(g.get('fs', g.get('font_size', 0)) for g in r)
        out.append((k, len(r), fs, ''.join(g['char'] for g in r)))
    return out


L = ['S505 b837 p4 transition (y 520-670)']
L.append('--- WORD ---')
for k, n, fs, t in lines(W, 'y', 520, 670):
    L.append('  y=%6.1f n=%2d fs=%.0f  %s' % (k, n, fs, t[:40]))
L.append('--- OXI ---')
for k, n, fs, t in lines(O, 'baseline', 520, 670):
    L.append('  y=%6.1f n=%2d fs=%.0f  %s' % (k, n, fs, t[:40]))
with io.open('c:/tmp/_s505_b837_out.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('wrote c:/tmp/_s505_b837_out.txt')
