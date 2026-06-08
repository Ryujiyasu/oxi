# -*- coding: utf-8 -*-
"""S510 line-pitch compare: Word vs Oxi cumulative line baselines + per-line pitch set.
Usage: python _s510_pitch.py <word.json> <oxi.json>  cp932-safe."""
import json, io, sys, statistics


def bl(p, k):
    g = json.load(io.open(p, encoding='utf-8'))['pages'][0]['glyphs']
    rows = {}
    for x in g:
        if x['char'].strip():
            rows.setdefault(round(x[k], 1), []).append(x)
    ys = sorted(rows); L = []
    for y in ys:
        if L and y - L[-1] < 4:
            continue
        L.append(y)
    return L


W = bl(sys.argv[1], 'y')
O = bl(sys.argv[2], 'baseline')
wp = [round(W[i + 1] - W[i], 2) for i in range(len(W) - 1) if 8 < W[i + 1] - W[i] < 20]
op = [round(O[i + 1] - O[i], 2) for i in range(len(O) - 1) if 8 < O[i + 1] - O[i] < 20]
print('Word   pitch set %-28s mean %.3f' % (str(sorted(set(wp))[:8]), statistics.mean(wp)))
print('Oxi    pitch set %-28s mean %.3f' % (str(sorted(set(op))[:8]), statistics.mean(op)))
print('first Oxi-Word = %.2f   last Oxi-Word = %.2f' % (O[0] - W[0], O[-1] - W[-1]))
