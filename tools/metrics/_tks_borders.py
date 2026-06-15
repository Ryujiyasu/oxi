# -*- coding: utf-8 -*-
"""tokyoshugyo cell-height localization: per-page horizontal table-border Y
positions, Word PDF (fitz) vs Oxi --dump-layout 'border' elements. On pages
where pagination agrees (delta 0), a divergence in border Y = table row-height
error (the doc-wide -1 drift hypothesis). Reports paired borders + row gaps."""
import os, sys, json, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

PDF  = os.path.join(tempfile.gettempdir(), 'tks_truth.pdf')
DUMP = os.path.join(tempfile.gettempdir(), 'tks_oxi_dump.json')

import fitz
doc = fitz.open(PDF)

def word_hlines(pi):
    pg = doc[pi]
    ys = []
    for dr in pg.get_drawings():
        for it in dr['items']:
            if it[0] == 'l' and abs(it[1].y - it[2].y) < 0.5:
                ys.append(it[1].y)
            elif it[0] == 're':
                ys.append(it[1].y0); ys.append(it[1].y1)
    # cluster within 1pt (border thickness pairs)
    ys = sorted(ys)
    out = []
    for y in ys:
        if not out or y - out[-1] > 1.2:
            out.append(y)
    return [round(y, 1) for y in out]

d = json.load(open(DUMP, encoding='utf-8'))
def oxi_hlines(pi):
    pg = d['pages'][pi]
    ys = sorted({round(el['y'], 1) for el in pg.get('elements', []) if el.get('type') == 'border'})
    out = []
    for y in ys:
        if not out or y - out[-1] > 1.2:
            out.append(y)
    return out

pages = sys.argv[1:] or list(range(1, 12))
for p in [int(x) for x in pages]:
    wp = word_hlines(p-1)
    op = oxi_hlines(p-1)
    print(f"\n=== page {p} ===  Word {len(wp)} borders, Oxi {len(op)} borders")
    print(f"  Word Y: {wp}")
    print(f"  Oxi  Y: {op}")
    # pair by nearest, show diff
    n = min(len(wp), len(op))
    if n:
        print("  paired (Word -> Oxi, dY):")
        for i in range(n):
            print(f"    {wp[i]:7.1f} -> {op[i]:7.1f}  dY={op[i]-wp[i]:+.1f}")
