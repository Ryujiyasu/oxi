# -*- coding: utf-8 -*-
"""S523 gate-truth instrument: per-LINE vertical drift profile from the ACTUAL gate pixels
(word_png vs oxi_png, 150dpi, the exact SSIM comparison). Builds a row-ink-density profile for
each image, finds text-line bands in the Word image, and for each band cross-correlates a local
window against the Oxi image to get the per-line vertical offset (px -> pt). This shows whether
the first-line offset is a UNIFORM SHIFT (a first-line fix would help the whole page) or DRIFTS/
COMPENSATES downstream (why a uniform first-line fix goes mixed, S494b). Renders fresh oxi_png
with the current (clean) binary. cp932-safe: UTF-8 file, ASCII out."""
import os, sys, subprocess, io, glob
import numpy as np
from PIL import Image
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
WORD_PNG = os.path.join(ROOT, 'pipeline_data', 'word_png')
DOCX = os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx')
DPI = 150
PT = DPI / 72.0  # px per pt

def render_oxi(docx, stem):
    out = os.path.join('c:/tmp', 's523_' + stem[:12])
    subprocess.run([EXE, os.path.abspath(docx), out, str(DPI)], capture_output=True, text=True)
    p = out + '_p1.png'
    return p if os.path.exists(p) else None

def row_ink(png):
    im = np.asarray(Image.open(png).convert('L'), dtype=np.float32)
    ink = (255.0 - im).sum(axis=1)  # darkness per row
    return ink, im.shape

def line_bands(ink, thresh=None):
    if thresh is None:
        thresh = ink.max() * 0.04
    bands = []; in_b = False; s = 0
    for i, v in enumerate(ink):
        if v > thresh and not in_b:
            in_b = True; s = i
        elif v <= thresh and in_b:
            in_b = False
            if i - s >= 3:
                bands.append((s, i))
    if in_b:
        bands.append((s, len(ink)))
    return bands

def best_shift(w_ink, o_ink, center, half=18):
    # cross-correlate o_ink vs w_ink near `center` over a +-half window of shifts
    lo = max(0, center - 40); hi = min(len(w_ink), center + 40)
    wseg = w_ink[lo:hi]
    best = 0; bestc = -1e18
    for sh in range(-half, half + 1):
        a = lo + sh; b = hi + sh
        if a < 0 or b > len(o_ink):
            continue
        oseg = o_ink[a:b]
        c = float((wseg * oseg).sum())
        if c > bestc:
            bestc = c; best = sh
    return best

def main():
    stems = sys.argv[1:]
    L = ['S523 per-line pixel drift (oxi_png - word_png, 150dpi gate pixels). +shift = Oxi LOWER than Word.']
    for stem in stems:
        g = glob.glob(os.path.join(DOCX, stem + '*.docx'))
        wdir = glob.glob(os.path.join(WORD_PNG, stem + '*'))
        if not g or not wdir:
            L.append('%s: missing docx or word_png' % stem); continue
        wpng = os.path.join(wdir[0], 'page_0001.png')
        if not os.path.exists(wpng):
            L.append('%s: no word page_0001.png' % stem); continue
        opng = render_oxi(g[0], stem)
        if not opng:
            L.append('%s: oxi render failed' % stem); continue
        w_ink, _ = row_ink(wpng); o_ink, _ = row_ink(opng)
        bands = line_bands(w_ink)
        L.append('--- %s : %d word text-bands' % (stem[:20], len(bands)))
        prof = []
        for (s, e) in bands[:24]:
            c = (s + e) // 2
            sh = best_shift(w_ink, o_ink, c)
            prof.append((c, sh))
        # report first few + drift summary
        for i, (c, sh) in enumerate(prof[:8]):
            L.append('   band%2d  word_row=%4d  shift=%+3dpx (%+.2fpt)' % (i, c, sh, sh / PT))
        if len(prof) >= 4:
            first = np.mean([p[1] for p in prof[:2]])
            last = np.mean([p[1] for p in prof[-2:]])
            L.append('   FIRST~%.1fpx (%.2fpt)  LAST~%.1fpx (%.2fpt)  drift=%.1fpx (%.2fpt) over page'
                     % (first, first / PT, last, last / PT, last - first, (last - first) / PT))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s523_out.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)

if __name__ == '__main__':
    main()
