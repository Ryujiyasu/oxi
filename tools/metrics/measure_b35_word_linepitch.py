# -*- coding: utf-8 -*-
"""S492u-2 — decisive COM-free test of the b35 snapToGrid=false hypothesis.
b35 has 13 <w:snapToGrid val=0> paras + docGrid linesAndChars linePitch=350 (17.5pt).
Oxi renders ~5 lines at 18.0pt (natural) vs the 17.5pt grid; the rest at 17.5.
Q: does WORD grid-snap the opt-out paras to 17.5 (=> Oxi's 18.0 is the bug, drifts below),
or does Word also use ~18.0 natural (=> drift is a natural-height mismatch elsewhere)?
Extract line-center Y from BOTH PNGs via ink-row projection peak-finding, histogram the
consecutive spacings (px->pt @150dpi: 1pt=2.083px). If Word has NO ~18.0 spacings but Oxi
does, Word snaps the opt-out paras. cp932-safe (UTF-8, ASCII out)."""
import json
import numpy as np
from PIL import Image
from scipy.signal import find_peaks

PT = 150.0 / 72.0  # px per pt @150dpi = 2.083

targets = json.load(open('c:/tmp/bottomN/targets.json', encoding='utf-8'))
t = next(x for x in targets if x['doc_id'] == 'b35123fe8efc' and x['page'] == 1)


def line_spacings(png, size=None):
    im = Image.open(png).convert('L')
    if size and im.size != size:
        im = im.resize(size)
    a = 255 - np.array(im, dtype=np.float32)  # ink high
    row = a.sum(axis=1)
    # smooth lightly, find line-center peaks (min distance ~10px = 4.8pt)
    thr = row.max() * 0.12
    peaks, _ = find_peaks(row, height=thr, distance=10)
    ys = peaks.astype(float)
    sp = np.diff(ys) / PT  # pt
    return im.size, ys, sp


wsz, wy, wsp = line_spacings(t['word_png'])
osz, oy, osp = line_spacings(t['oxi_png'], size=wsz)


def hist(sp, lo=14.0, hi=22.0):
    # focus on the body-text line-spacing band (ignore big para gaps & noise)
    s = sp[(sp >= lo) & (sp <= hi)]
    bins = {}
    for v in s:
        k = round(v * 2) / 2  # 0.5pt bins
        bins[k] = bins.get(k, 0) + 1
    return dict(sorted(bins.items()))


print("b35 p1  word_png peaks=%d  oxi peaks=%d  (size=%s)" % (len(wy), len(oy), wsz))
print("grid linePitch=350tw=17.5pt; opt-out paras (snapToGrid=0)=13 doc-wide")
print("\nWORD line-spacing histogram (0.5pt bins, 14-22pt band):")
for k, v in hist(wsp).items():
    bar = '#' * v
    mark = ' <-17.5 GRID' if abs(k - 17.5) < 0.26 else (' <-18.0 (natural/over)' if abs(k - 18.0) < 0.26 else '')
    print("  %5.1fpt x%-3d %s%s" % (k, v, bar, mark))
print("\nOXI line-spacing histogram:")
for k, v in hist(osp).items():
    bar = '#' * v
    mark = ' <-17.5 GRID' if abs(k - 17.5) < 0.26 else (' <-18.0 (natural/over)' if abs(k - 18.0) < 0.26 else '')
    print("  %5.1fpt x%-3d %s%s" % (k, v, bar, mark))

w18 = ((wsp >= 17.75) & (wsp <= 18.25)).sum()
o18 = ((osp >= 17.75) & (osp <= 18.25)).sum()
w175 = ((wsp >= 17.25) & (wsp < 17.75)).sum()
o175 = ((osp >= 17.25) & (osp < 17.75)).sum()
print("\n=== VERDICT ===")
print("Word  17.5pt-band x%d, 18.0pt-band x%d" % (w175, w18))
print("Oxi   17.5pt-band x%d, 18.0pt-band x%d" % (o175, o18))
if w18 <= 1 and o18 >= 3:
    print(">>> WORD GRID-SNAPS opt-out paras to 17.5; Oxi's 18.0 natural is the BUG (drifts below).")
    print(">>> FIX: in linesAndChars docGrid, snap line height to linePitch EVEN when snapToGrid=false.")
elif w18 >= 3:
    print(">>> Word ALSO has ~18.0 lines => not a grid-snap bug; natural-height mismatch elsewhere.")
else:
    print(">>> Inconclusive from PNG peaks; needs COM per-line Information(6) probe.")
