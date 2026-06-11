# -*- coding: utf-8 -*-
"""S541b: char-aligned advance diff between Word raw chars (c:/tmp/_s541_word_raw.txt
from _s541_page_charcmp.py) and the Oxi dump (c:/tmp/s541_dump.json).
For each Word line, align chars to Oxi elements by text and report per-char
width deltas > tol. Usage: python _s541_advdiff.py [tol=0.3]
"""
import io, json, sys
from collections import defaultdict

tol = float(sys.argv[1]) if len(sys.argv) > 1 else 0.3

# --- load word chars grouped by para/line
wparas = []
cur = None
for line in io.open('c:/tmp/_s541_word_raw.txt', encoding='utf-8'):
    line = line.rstrip('\n')
    if line.startswith('PARA|'):
        cur = {'text': line.split('|', 2)[2], 'lines': defaultdict(list)}
        wparas.append(cur)
    elif line.startswith('C|') and cur is not None:
        _, i, y, x, ch = line.split('|', 4)
        cur['lines'][float(y)].append((int(i), float(x), ch))

# --- load oxi paras: flatten to per-char (single-char elements only; multi-char split evenly)
d = json.load(io.open('c:/tmp/s541_dump.json', encoding='utf-8'))
oxi_paras = {}
for pi, page in enumerate(d['pages'][:3]):
    groups = defaultdict(list)
    for e in page.get('elements', []):
        if e.get('type') == 'text':
            groups[(e.get('para_idx'), e.get('cell_para_idx'), e.get('cell_row_idx'), e.get('cell_col_idx'))].append(e)
    for k, els in groups.items():
        lines = defaultdict(list)
        for e in els:
            lines[round(e['y'], 1)].append(e)
        chars = []  # (ch, w) in reading order
        for y in sorted(lines):
            row = sorted(lines[y], key=lambda e: e.get('x', 0))
            for e in row:
                t = e.get('text', '')
                if not t:
                    continue
                w = e.get('w', 0)
                for ch in t:
                    chars.append((ch, w / len(t)))
        if chars:
            txt = ''.join(c for c, _ in chars)
            oxi_paras[txt[:10]] = (chars, txt)

out = io.open('c:/tmp/s541_advdiff.txt', 'w', encoding='utf-8')
for wp in wparas:
    # word chars in para order with advances (per line; last char of line gets None)
    wchars = []
    for y in sorted(wp['lines']):
        chs = sorted(wp['lines'][y])
        for k in range(len(chs)):
            adv = chs[k + 1][1] - chs[k][1] if k + 1 < len(chs) else None
            wchars.append((chs[k][2], adv))
    wtxt = ''.join(c for c, _ in wchars)
    if len(wtxt) < 4:
        continue
    # find oxi para by prefix
    hit = None
    for pref, (chars, otxt) in oxi_paras.items():
        if otxt[:8] == wtxt[:8] or otxt[:6] == wtxt[:6]:
            hit = (chars, otxt)
            break
    if not hit:
        out.write('NOMATCH %s\n' % wtxt[:16])
        continue
    ochars, otxt = hit
    if otxt[:40] != wtxt[:40]:
        # texts diverge (different break artifacts ok, chars must align)
        pass
    n = min(len(wchars), len(ochars))
    diffs = []
    for i in range(n):
        wc, wadv = wchars[i]
        oc, ow = ochars[i]
        if wc != oc:
            diffs.append((i, 'TEXTDIFF %r vs %r' % (wc, oc)))
            break
        if wadv is None:
            continue
        if abs(wadv - ow) > tol:
            diffs.append((i, '%s W=%.2f O=%.2f d=%+.2f' % (wc, wadv, ow, ow - wadv)))
    if diffs:
        out.write('=== %s (W %d ch / O %d ch) ===\n' % (wtxt[:20], len(wchars), len(ochars)))
        for i, msg in diffs[:25]:
            out.write('  [%3d] %s\n' % (i, msg))
out.close()
print('ok -> c:/tmp/s541_advdiff.txt')
