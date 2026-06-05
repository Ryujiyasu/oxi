# -*- coding: utf-8 -*-
"""Dump the ACTUAL text (UTF-8) of baseline-clustered lines in a y-range, Word and Oxi,
with each line's first-x, end-x (last glyph x + approx advance), and char count.
Output is UTF-8 to a file (Read it; do NOT eyeball console on cp932)."""
import json, sys, statistics


def cluster(glyphs, ykey):
    gs = sorted([g for g in glyphs if g['char'].strip()], key=lambda g: (round(g[ykey], 1), g['x']))
    lines = []
    for g in gs:
        if lines and abs(g[ykey] - lines[-1]['y']) < 4.0:
            lines[-1]['gs'].append(g)
        else:
            lines.append({'y': g[ykey], 'gs': [g]})
    for L in lines:
        L['gs'].sort(key=lambda g: g['x'])
        L['y'] = statistics.median(g[ykey] for g in L['gs'])
    return lines


def main():
    wpath, opath, pidx, ylo, yhi = sys.argv[1], sys.argv[2], int(sys.argv[3]), float(sys.argv[4]), float(sys.argv[5])
    out = sys.argv[6] if len(sys.argv) > 6 else 'c:/tmp/_s496_linechars.txt'
    W = json.load(open(wpath, encoding='utf-8'))['pages'][pidx]['glyphs']
    O = json.load(open(opath, encoding='utf-8'))['pages'][pidx]['glyphs']

    def emit(f, lines, label):
        f.write('=== %s p%d y[%g,%g] ===\n' % (label, pidx, ylo, yhi))
        for L in [L for L in lines if ylo <= L['y'] <= yhi]:
            gs = L['gs']
            txt = ''.join(g['char'] for g in gs)
            x0 = gs[0]['x']; xlast = gs[-1]['x']
            f.write('y=%6.1f n=%2d x0=%6.1f xlast=%6.1f  |%s|\n' % (L['y'], len(gs), x0, xlast, txt))

    with open(out, 'w', encoding='utf-8') as f:
        emit(f, cluster(W, 'y'), 'WORD')
        f.write('\n')
        emit(f, cluster(O, 'baseline'), 'OXI')
    print('wrote', out)


if __name__ == '__main__':
    main()
