# -*- coding: utf-8 -*-
"""For each long line on a page, report Word vs Oxi mean CJK char advance (matched).
Reveals whether Oxi's per-char advance is a consistent ratio too wide (systematic
char-width/grid-pitch bug) or line-specific. cp932-safe."""
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


def main():
    wpath, opath, pidx = sys.argv[1], sys.argv[2], int(sys.argv[3])
    minn = int(sys.argv[4]) if len(sys.argv) > 4 else 25
    out = sys.argv[5] if len(sys.argv) > 5 else 'c:/tmp/_s496_adv.txt'
    W = json.load(open(wpath, encoding='utf-8'))['pages'][pidx]['glyphs']
    O = json.load(open(opath, encoding='utf-8'))['pages'][pidx]['glyphs']
    WL = cluster(W, 'y'); OL = cluster(O, 'baseline')
    rows = []
    for wl in WL:
        if len(wl['gs']) < minn:
            continue
        ol = min(OL, key=lambda L: abs(L['y'] - wl['y']))
        wc = [g['char'] for g in wl['gs']]; oc = [g['char'] for g in ol['gs']]
        sm = difflib.SequenceMatcher(None, wc, oc, autojunk=False)
        m = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == 'equal':
                for d in range(i2 - i1):
                    m.append((wl['gs'][i1 + d], ol['gs'][j1 + d]))
        if len(m) < minn:
            continue
        wadv = (m[-1][0]['x'] - m[0][0]['x']) / (len(m) - 1)
        oadv = (m[-1][1]['x'] - m[0][1]['x']) / (len(m) - 1)
        fs = statistics.median(g.get('font_size', g.get('fs', 0)) for _, g in m)
        rows.append((wl['y'], len(m), fs, wadv, oadv, oadv - wadv, oadv / wadv if wadv else 0))
    with open(out, 'w', encoding='utf-8') as f:
        f.write('   y    n   fs   Wadv   Oadv   dAdv   O/W ratio\n')
        for y, n, fs, wa, oa, d, r in rows:
            f.write('%6.1f %3d %4.1f %6.3f %6.3f %+.3f  %.4f\n' % (y, n, fs, wa, oa, d, r))
    print('wrote', out)


if __name__ == '__main__':
    main()
