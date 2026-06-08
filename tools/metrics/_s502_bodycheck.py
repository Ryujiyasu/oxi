# -*- coding: utf-8 -*-
"""S502 body-path check: compare Word vs Oxi (current) first-char x + advance for a body
centered grid line, to see if the body alignment path has the same grid-natural-width bug
as the cell path (S502 is cell-only). cp932-safe: UTF-8 file, ASCII out.
Usage: python _s502_bodycheck.py <w.json> <oxi.json> <page> <needle>"""
import json, io, sys

wj, oj, pidx, needle = sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4]
W = json.load(io.open(wj, encoding='utf-8'))['pages'][pidx]['glyphs']
O = json.load(io.open(oj, encoding='utf-8'))['pages'][pidx]['glyphs']


def find(glyphs, getx):
    chars = [g['char'] for g in glyphs]
    for i in range(len(chars) - len(needle) + 1):
        if ''.join(chars[i:i + len(needle)]) == needle:
            run = glyphs[i:i + len(needle)]
            xs = [getx(g) for g in run]
            advs = [xs[k + 1] - xs[k] for k in range(len(xs) - 1)]
            return xs[0], xs[-1], (sum(advs) / len(advs) if advs else 0)
    return None, None, None


wf, wl, wa = find(W, lambda g: g['x'])
of, ol, oa = find(O, lambda g: g['x'])
L = ['S502 body-path check  needle=%s  page=%d' % (needle, pidx)]
if wf is None or of is None:
    L.append('NOT FOUND: word=%s oxi=%s' % (wf is not None, of is not None))
else:
    L.append('         first_x   last_x   mean_adv')
    L.append('WORD    %8.2f %8.2f %8.3f' % (wf, wl, wa))
    L.append('OXI     %8.2f %8.2f %8.3f' % (of, ol, oa))
    L.append('first-char dx (OXI-WORD) = %+.2f  (line width: WORD %.2f OXI %.2f)' % (
        of - wf, wl - wf, ol - of))
with io.open('c:/tmp/_s502_bodycheck_out.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('wrote c:/tmp/_s502_bodycheck_out.txt')
