# -*- coding: utf-8 -*-
"""Horizontal per-line profile: Word PDF glyphs vs Oxi dump-glyphs. For each matched line
print left-edge x (first glyph), right-edge x (last glyph), and glyph count, Word vs Oxi.
Localizes a horizontal block offset to left-shift / right-shift / justify-spread. cp932-safe."""
import json, sys, statistics, difflib
K = 0.859


def cluster(glyphs, getby, tol=4.0):
    gs = sorted(glyphs, key=lambda g: (round(getby(g), 1), g['x']))
    lines = []
    for g in gs:
        by = getby(g)
        if lines and abs(by - lines[-1]['by']) < tol:
            lines[-1]['gs'].append(g)
        else:
            lines.append({'by': by, 'gs': [g]})
    for L in lines:
        L['gs'].sort(key=lambda g: g['x'])
        L['by'] = statistics.median(getby(x) for x in L['gs'])
    return lines


def main():
    wpath, opath, pidx = sys.argv[1], sys.argv[2], int(sys.argv[3])
    W = json.load(open(wpath, encoding='utf-8'))['pages'][pidx]['glyphs']
    O = json.load(open(opath, encoding='utf-8'))['pages'][pidx]['glyphs']
    wl = cluster([g for g in W if g['char'].strip()], lambda g: g['y'])
    ol = cluster([g for g in O if g['char'].strip()], lambda g: g['top'] + g['font_size'] * K)
    # match lines by order (both sorted by y); print where counts match
    print(' Wy    Oy    | wL    oL   dL  | wR    oR   dR  | wn on  text')
    n = min(len(wl), len(ol))
    for i in range(n):
        w, o = wl[i], ol[i]
        wL, wR = w['gs'][0]['x'], w['gs'][-1]['x']
        oL, oR = o['gs'][0]['x'], o['gs'][-1]['x']
        wn, on = len(w['gs']), len(o['gs'])
        # crude text sample (ascii only)
        samp = ''.join(c['char'] if ord(c['char']) < 128 else '.' for c in w['gs'][:10])
        print(' %5.1f %5.1f | %5.1f %5.1f %+5.1f | %5.1f %5.1f %+5.1f | %2d %2d  %s'
              % (w['by'], o['by'], wL, oL, oL - wL, wR, oR, oR - wR, wn, on, samp))


if __name__ == '__main__':
    main()
