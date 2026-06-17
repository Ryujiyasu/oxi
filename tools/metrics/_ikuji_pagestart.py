# -*- coding: utf-8 -*-
"""For each page, show Word's first content line vs Oxi's first content line.
If Oxi has an EXTRA leading (partial) line that Word fit at the bottom of the
previous page, the page inherited a 1-line downward shift. This localizes where
the wi355/wi440 shift ORIGINATES and how far back it cascades."""
import os, sys, json, tempfile, difflib
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz

PDF  = os.path.join(tempfile.gettempdir(), 'ikd_truth.pdf')
DUMP = os.path.join(tempfile.gettempdir(), 'ikuji_dump.json')


def word_first_last(pno):
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
                    if b[1] < 65 or b[1] > 795:   # strip header + footer (- N -)
                        continue
                    chars.append((ch['c'], b[0], b[1], b[3], (b[1] + b[3]) / 2))
    if not chars:
        return None
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
        out.append((round(min(t[2] for t in r), 1), round(max(t[3] for t in r), 1),
                    ''.join(t[0] for t in r)))
    return out


def oxi_first_last(pno):
    d = json.load(open(DUMP, encoding='utf-8'))
    for pg in d['pages']:
        if pg['page'] != pno:
            continue
        rows = {}
        for el in pg.get('elements', []):
            if el.get('type') != 'text' or not el.get('text'):
                continue
            y = round(el['y'], 1)
            if y < 60 or y > 795:   # strip header + footer
                continue
            rows.setdefault(y, []).append(el)
        out = []
        for key in sorted(rows):
            row = sorted(rows[key], key=lambda e: e['x'])
            out.append((key, round(key + max(e.get('h', 0) for e in row), 1),
                        ''.join(e.get('text', '') for e in row),
                        row[0].get('para_idx'), row[0].get('char_offset')))
        return out
    return []


npages = len(fitz.open(PDF))
print(f"pg | Wln | Oln | Word first content line              | Oxi first content line (para/off)")
for p in range(1, npages + 1):
    W = word_first_last(p) or []
    O = oxi_first_last(p) or []
    wf = W[0] if W else ('', '', '(none)')
    of = O[0] if O else ('', '', '(none)', '', '')
    # do the first lines match (same leading text)?
    wtxt = wf[2].replace(' ', '').replace('　', '')[:18]
    otxt = of[2].replace(' ', '').replace('　', '')[:18]
    match = '  ' if wtxt[:8] == otxt[:8] else '!='
    print(f"p{p:2d}| {len(W):3d} | {len(O):3d} | {wtxt:<20} y0={wf[0]:6} {match} {otxt:<20} y0={of[0]:6} p{of[3]}/{of[4]}")
    # also show last content line of each
    wl = W[-1] if W else ('', '', '')
    ol = O[-1] if O else ('', '', '', '', '')
    print(f"    last: W '{wl[2][:22]:<22}' y0={wl[0]:6} y1={wl[1]:6}  |  O '{ol[2][:22]:<22}' y0={ol[0]:6} y1={ol[1]:6}")
