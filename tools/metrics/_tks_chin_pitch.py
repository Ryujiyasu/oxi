# -*- coding: utf-8 -*-
"""Pin tokyoshugyo #2 (賃金 chapter page-bottom over-fit): is Oxi's cell line
PITCH smaller, its page-bottom more lenient, or its page-start higher? Compare
the 賃金 chapter cell line Y positions Oxi (dump) vs Word (PDF), per page.

The 賃金 chapter: Word p46-64, Oxi p47-64 (shifted +1 by #1). Compare Oxi page
N+1 to Word page N. For each, list cell-text line top-Ys → pitch (consecutive
gap) + first/last line Y + count. If Oxi pitch < Word pitch → Oxi packs tighter
(line-height bug). If pitch == but Oxi has more lines reaching lower → page-bottom
leniency. If Oxi first-Y higher → page-start."""
import os, sys, json, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz

PDF = os.path.join(tempfile.gettempdir(), 'tks_truth.pdf')
DUMP = os.path.join(tempfile.gettempdir(), 'tks_base.json')


def word_page_line_ys(pno):
    pg = fitz.open(PDF)[pno-1]
    rd = pg.get_text('rawdict')
    ys = []
    for blk in rd['blocks']:
        if blk.get('type', 0) != 0:
            continue
        for ln in blk.get('lines', []):
            txt = ''.join(c['c'] for sp in ln['spans'] for c in sp['chars']).strip()
            if not txt:
                continue
            y0 = min(c['bbox'][1] for sp in ln['spans'] for c in sp['chars'])
            if y0 < 60 or y0 > 800:
                continue
            ys.append(round(y0, 1))
    return sorted(set(ys))


def oxi_page_line_ys(pno):
    d = json.load(open(DUMP, encoding='utf-8'))
    for pg in d['pages']:
        if pg['page'] != pno:
            continue
        ys = set()
        for el in pg.get('elements', []):
            if el.get('type') == 'text' and el.get('text') and 60 < el['y'] < 800:
                ys.add(round(el['y'], 1))
        return sorted(ys)
    return []


def pitch_stats(ys):
    if len(ys) < 2:
        return (len(ys), None, None, None)
    gaps = [round(ys[i+1]-ys[i], 2) for i in range(len(ys)-1)]
    from collections import Counter
    common = Counter(g for g in gaps if g < 30).most_common(1)
    mode = common[0][0] if common else None
    return (len(ys), ys[0], ys[-1], mode)


# Find the 賃金 chapter start page (look for 賃金 in a heading-ish line)
d = json.load(open(DUMP, encoding='utf-8'))
chin_oxi = None
for pg in d['pages']:
    txt = ''.join(e.get('text', '') for e in pg.get('elements', []) if e.get('type') == 'text')
    if '賃金' in txt and ('賃金規程' in txt or '第' in txt and '章' in txt and '賃金' in txt):
        chin_oxi = pg['page']
        break
print(f"賃金 chapter Oxi start ~p{chin_oxi}")

print(f"\n{'pair':>10} | {'Oxi: n  first  last  pitch':>30} | {'Word: n  first  last  pitch':>30}")
# compare Oxi p(N+1) to Word pN for the 賃金 chapter (memory: Word p46-64 / Oxi p47-64)
for wp in range(46, 65):
    op = wp + 1
    on, of, ol, opi = pitch_stats(oxi_page_line_ys(op))
    wn, wf, wl, wpi = pitch_stats(word_page_line_ys(wp))
    print(f"  Wp{wp} Op{op} | O: {on:3d} {str(of):>6} {str(ol):>6} {str(opi):>6} | "
          f"W: {wn:3d} {str(wf):>6} {str(wl):>6} {str(wpi):>6} | dN={on-wn:+d}")
