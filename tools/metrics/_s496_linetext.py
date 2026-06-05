# -*- coding: utf-8 -*-
"""Dump per-line (baseline-clustered) y + ascii-prefix + codepoints for Word and Oxi
in a y-range on one page. cp932-safe: ascii + U+codes only to a file."""
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


def fmt(L):
    txt = ''.join(g['char'] for g in L['gs'])
    asc = ''.join(c if 32 <= ord(c) < 128 else '.' for c in txt)[:30]
    fs = statistics.median(g.get('font_size', g.get('fs', 0)) for g in L['gs'])
    return '%6.1f n=%2d fs=%4.1f x0=%6.1f  %s' % (L['y'], len(L['gs']), fs, L['gs'][0]['x'], asc)


def main():
    wpath, opath, pidx, ylo, yhi = sys.argv[1], sys.argv[2], int(sys.argv[3]), float(sys.argv[4]), float(sys.argv[5])
    out = sys.argv[6] if len(sys.argv) > 6 else 'c:/tmp/_s496_linetext.txt'
    W = json.load(open(wpath, encoding='utf-8'))['pages'][pidx]['glyphs']
    O = json.load(open(opath, encoding='utf-8'))['pages'][pidx]['glyphs']
    wl = [L for L in cluster(W, 'y') if ylo <= L['y'] <= yhi]
    ol = [L for L in cluster(O, 'baseline') if ylo <= L['y'] <= yhi]
    with open(out, 'w', encoding='utf-8') as f:
        f.write('=== WORD p%d y[%g,%g] ===\n' % (pidx, ylo, yhi))
        prev = None
        for L in wl:
            gap = '' if prev is None else '  (+%.1f)' % (L['y'] - prev)
            f.write(fmt(L) + gap + '\n'); prev = L['y']
        f.write('\n=== OXI p%d y[%g,%g] ===\n' % (pidx, ylo, yhi))
        prev = None
        for L in ol:
            gap = '' if prev is None else '  (+%.1f)' % (L['y'] - prev)
            f.write(fmt(L) + gap + '\n'); prev = L['y']
    print('wrote', out)


if __name__ == '__main__':
    main()
