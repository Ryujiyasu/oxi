# -*- coding: utf-8 -*-
"""Compare Word-PDF vs Oxi per-line Y on pages 13 and 16 (the wi355/wi440 spill
pages). Pin: does Word position its bottom lines HIGHER (line-height drift) so
the spill line fits, or is Word's page-bottom limit genuinely more lenient than
Oxi's? Prints, for each page, the char-aligned Word-line-y vs Oxi-line-y and the
last/spill line geometry."""
import os, sys, json, tempfile, difflib
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz

PDF  = os.path.join(tempfile.gettempdir(), 'ikd_truth.pdf')
DUMP = os.environ.get('IKD_DUMP', os.path.join(tempfile.gettempdir(), 'ikuji_dump.json'))


def word_page_lines(pno):
    """Word PDF page pno -> list of (y0_top, y1_bottom, text)."""
    doc = fitz.open(PDF)
    pg = doc[pno - 1]
    rd = pg.get_text('rawdict')
    chars = []
    for blk in rd['blocks']:
        if blk.get('type', 0) != 0:
            continue
        for ln in blk.get('lines', []):
            for sp in ln['spans']:
                for ch in sp['chars']:
                    b = ch['bbox']
                    if b[1] < 65 or b[1] > 820:
                        continue
                    chars.append((ch['c'], b[0], b[1], b[3], (b[1] + b[3]) / 2))
    chars.sort(key=lambda t: (t[4], t[1]))
    rows, cur, cy = [], [], None
    for c, x0, y0, y1, yc in chars:
        if cy is None or abs(yc - cy) <= 7.0:
            cur.append((c, x0, y0, y1, yc)); cy = yc if cy is None else (cy*(len(cur)-1)+yc)/len(cur)
        else:
            rows.append(cur); cur = [(c, x0, y0, y1, yc)]; cy = yc
    if cur:
        rows.append(cur)
    out = []
    for r in rows:
        r.sort(key=lambda t: t[1])
        y0 = min(t[2] for t in r); y1 = max(t[3] for t in r)
        out.append((round(y0, 2), round(y1, 2), ''.join(t[0] for t in r)))
    return out


def oxi_page_lines(pno):
    d = json.load(open(DUMP, encoding='utf-8'))
    for pg in d['pages']:
        if pg['page'] != pno:
            continue
        rows = {}
        for el in pg.get('elements', []):
            if el.get('type') != 'text' or not el.get('text'):
                continue
            rows.setdefault(round(el['y'], 1), []).append(el)
        out = []
        for key in sorted(rows):
            row = sorted(rows[key], key=lambda e: e['x'])
            txt = ''.join(e.get('text', '') for e in row)
            h = max(e.get('h', 0) for e in row)
            tyo = row[0].get('text_y_off', 0)
            out.append((round(key, 2), round(key + h, 2), tyo, txt))
        return out
    return []


# Page geometry from the dump
d = json.load(open(DUMP, encoding='utf-8'))
pg0 = d['pages'][0]
print("dump page keys:", sorted(k for k in pg0.keys() if k != 'elements'))
for k in pg0:
    if k != 'elements':
        print(f"   {k} = {pg0[k]}")

PAGES = (13, 16)
_a = [x for x in sys.argv[1:] if x.isdigit()]
if _a:
    PAGES = tuple(int(x) for x in _a)
for pno in PAGES:
    print(f"\n========== PAGE {pno} ==========")
    W = word_page_lines(pno)
    O = oxi_page_lines(pno)
    print(f"Word lines: {len(W)}   Oxi lines: {len(O)}")
    # char-stream align the two pages to pair lines
    wtext = ''.join(t[2].replace(' ', '').replace('　', '') for t in W)
    otext = ''.join(t[3].replace(' ', '').replace('　', '') for t in O)
    # build per-char line-id maps
    wmap, omap = [], []
    for li, t in enumerate(W):
        for c in t[2]:
            if c not in ' 　\t': wmap.append(li)
    for li, t in enumerate(O):
        for c in t[3]:
            if c not in ' 　\t': omap.append(li)
    sm = difflib.SequenceMatcher(None, wtext, otext, autojunk=False)
    pair = {}  # oxi_line -> word_line (dominant)
    from collections import Counter
    cnt = {}
    for blk in sm.get_matching_blocks():
        for k in range(blk.size):
            wl = wmap[blk.a + k]; ol = omap[blk.b + k]
            cnt.setdefault(ol, Counter())[wl] += 1
    for ol, c in cnt.items():
        pair[ol] = c.most_common(1)[0][0]
    print(f"  {'Oxi y0':>7} {'Oy1':>7} {'tyo':>4} | {'Word y0':>7} {'Wy1':>7} | dW0  | text")
    for ol, o in enumerate(O):
        wl = pair.get(ol)
        if wl is not None:
            w = W[wl]
            dW0 = o[0] - w[0]
            print(f"  {o[0]:7.2f} {o[1]:7.2f} {o[2]:4.1f} | {w[0]:7.2f} {w[1]:7.2f} | {dW0:+5.2f} | {o[3][:30]}")
        else:
            print(f"  {o[0]:7.2f} {o[1]:7.2f} {o[2]:4.1f} | {'—':>7} {'—':>7} |   —   | {o[3][:30]}")
    # the spill line = first Word line whose y0 is below Oxi's last line
    o_last_y0 = O[-1][0]
    o_last_y1 = O[-1][1]
    spill = [w for w in W if w[0] > o_last_y0 + 5]
    print(f"  Oxi last line: y0={o_last_y0:.2f} y1={o_last_y1:.2f}")
    if spill:
        s = spill[0]
        print(f"  *** Word SPILL line (Oxi pushed to next pg): y0={s[0]:.2f} y1={s[1]:.2f} '{s[2][:34]}'")
        print(f"      Word spill-line ink-height = {s[1]-s[0]:.2f}pt; Word last-line y1 = {W[-1][1]:.2f}")
