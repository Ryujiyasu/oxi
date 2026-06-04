# -*- coding: utf-8 -*-
"""Isolate first-line start offset vs line pitch. Word PDF glyphs (baseline y) and Oxi
dump-glyphs (top + fs*K). Print first 4 line baselines for both + pitch. cp932-safe."""
import json, sys, statistics
K = 0.859


def lines(glyphs, getby, tol=4.0):
    gs = sorted(glyphs, key=lambda g: getby(g))
    out = []
    for g in gs:
        by = getby(g)
        if out and abs(by - out[-1][0]) < tol:
            out[-1][1].append(g)
        else:
            out.append([by, [g]])
    return [(statistics.median(getby(x) for x in gg), len(gg)) for by, gg in out]


def main():
    wpath, opath, pidx = sys.argv[1], sys.argv[2], int(sys.argv[3])
    W = json.load(open(wpath, encoding='utf-8'))['pages'][pidx]['glyphs']
    O = json.load(open(opath, encoding='utf-8'))['pages'][pidx]['glyphs']
    wl = lines([g for g in W if g['char'].strip()], lambda g: g['y'])
    ol = lines([g for g in O if g['char'].strip()], lambda g: g['top'] + g['font_size'] * K)
    ofs = [g['font_size'] for g in O if g['char'].strip()]
    print('first font_size(Oxi):', ofs[0] if ofs else '?')
    print(' idx   word_by  oxi_by   d(oxi-word)   w_pitch  o_pitch')
    for i in range(min(6, len(wl), len(ol))):
        wby, wn = wl[i]; oby, on = ol[i]
        wp = wl[i][0] - wl[i-1][0] if i else 0
        op = ol[i][0] - ol[i-1][0] if i else 0
        print('  %2d   %7.2f %7.2f   %+7.2f      %6.2f  %6.2f' % (i, wby, oby, oby - wby, wp, op))


if __name__ == '__main__':
    main()
