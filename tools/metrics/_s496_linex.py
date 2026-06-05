# -*- coding: utf-8 -*-
"""Compare per-line START x (leftmost glyph) Word vs Oxi, matched by line content.
Shows which lines are horizontally shifted. cp932-safe ASCII out + first chars as repr."""
import json, sys, statistics, difflib


def cluster(glyphs, ykey):
    gs = sorted(glyphs, key=lambda g: (round(g[ykey], 1), g['x']))
    lines = []
    for g in gs:
        if lines and abs(g[ykey] - lines[-1]['y']) < 4.0:
            lines[-1]['gs'].append(g)
        else:
            lines.append({'y': g[ykey], 'gs': [g]})
    for L in lines:
        L['gs'].sort(key=lambda g: g['x'])
        L['y'] = statistics.median(g[ykey] for g in L['gs'])
        L['x0'] = L['gs'][0]['x']
        L['txt'] = ''.join(g['char'] for g in L['gs'])
    return lines


def main():
    wpath, opath, pidx = sys.argv[1], sys.argv[2], int(sys.argv[3])
    out = sys.argv[4] if len(sys.argv) > 4 else 'c:/tmp/_s496_linex.txt'
    W = [g for g in json.load(open(wpath, encoding='utf-8'))['pages'][pidx]['glyphs'] if g['char'].strip()]
    O = [g for g in json.load(open(opath, encoding='utf-8'))['pages'][pidx]['glyphs'] if g['char'].strip()]
    wl = cluster(W, 'y')
    ol = cluster(O, 'baseline')
    # match Word lines to Oxi lines by text via difflib on concatenated first-8 chars
    wtxt = [L['txt'][:12] for L in wl]
    otxt = [L['txt'][:12] for L in ol]
    sm = difflib.SequenceMatcher(None, wtxt, otxt, autojunk=False)
    rows = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            for d in range(i2 - i1):
                w = wl[i1 + d]; o = ol[j1 + d]
                rows.append((w['y'], w['x0'], o['x0'], o['x0'] - w['x0'], w['txt'][:16]))
    with open(out, 'w', encoding='utf-8') as f:
        f.write('page %d  matched_lines %d\n' % (pidx, len(rows)))
        f.write(' line_y   Wx0    Oxx0   dx     text\n')
        for y, wx, ox, dx, t in rows:
            f.write('  %6.1f %6.1f %6.1f %+6.2f  %s\n' % (y, wx, ox, dx, t))
        dxs = [r[3] for r in rows]
        if dxs:
            f.write('\ndx median %.2f  mean %.2f  min %.2f max %.2f\n'
                    % (statistics.median(dxs), statistics.mean(dxs), min(dxs), max(dxs)))
    print('wrote', out)


if __name__ == '__main__':
    main()
