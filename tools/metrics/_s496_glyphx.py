# -*- coding: utf-8 -*-
"""Per-glyph horizontal x profile for ONE line (matched Word<->Oxi by content order).
Shows whether interior-glyph x drift is a uniform shift (Lever B/weight) or a
per-glyph accumulation across the line (Lever A char-grid justify). cp932-safe."""
import json, sys, statistics, difflib


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


def pick(lines, y):
    return min(lines, key=lambda L: abs(L['y'] - y))


def main():
    wpath, opath, pidx, wy = sys.argv[1], sys.argv[2], int(sys.argv[3]), float(sys.argv[4])
    out = sys.argv[5] if len(sys.argv) > 5 else 'c:/tmp/_s496_glyphx.txt'
    W = json.load(open(wpath, encoding='utf-8'))['pages'][pidx]['glyphs']
    O = json.load(open(opath, encoding='utf-8'))['pages'][pidx]['glyphs']
    wl = pick(cluster(W, 'y'), wy)
    ol = pick(cluster(O, 'baseline'), wy)
    wc = [g['char'] for g in wl['gs']]; oc = [g['char'] for g in ol['gs']]
    sm = difflib.SequenceMatcher(None, wc, oc, autojunk=False)
    rows = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            for d in range(i2 - i1):
                w = wl['gs'][i1 + d]; o = ol['gs'][j1 + d]
                rows.append((w['x'], o['x'], o['x'] - w['x']))
    with open(out, 'w', encoding='utf-8') as f:
        f.write('Word line y=%.1f (n=%d) <-> Oxi y=%.1f (n=%d)  matched %d\n'
                % (wl['y'], len(wl['gs']), ol['y'], len(ol['gs']), len(rows)))
        if rows:
            cal = rows[0][2]
            f.write('  idx   Wx     Ox      dx    dx-first(interior accumulation)\n')
            for i, (wx, ox, dx) in enumerate(rows):
                f.write('  %3d %7.2f %7.2f %+6.2f  %+6.2f\n' % (i, wx, ox, dx, dx - cal))
            f.write('first dx %+.2f  last dx %+.2f  span %+.2f  (advance W %.2f O %.2f over %d gaps)\n'
                    % (rows[0][2], rows[-1][2], rows[-1][2] - rows[0][2],
                       (rows[-1][0] - rows[0][0]) / max(1, len(rows) - 1),
                       (rows[-1][1] - rows[0][1]) / max(1, len(rows) - 1), len(rows) - 1))
    print('wrote', out)


if __name__ == '__main__':
    main()
