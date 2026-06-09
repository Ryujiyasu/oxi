# -*- coding: utf-8 -*-
"""S531 gate-truth HORIZONTAL drift instrument (the horizontal analog of _s523_pixel_drift).

S523 measured only VERTICAL per-line drift and found grid=none docs (683f/e3c545) ~0 -> labeled
their gap "weight/AA". But S509 hypothesized a HORIZONTAL justify-snap drift (Word device-snaps
each char advance in discrete bumps; Oxi distributes uniformly) -> cumulative right-ward drift on
each justified line. S523 could NOT see that (vertical only). This tool tests it in the gate pixels.

For each Word text-line band, slide a narrow window across the line and cross-correlate the Word
column-ink profile vs the Oxi one over a +-8px search (smaller than the ~22px CJK glyph pitch to
avoid period-ambiguity). Record best horizontal shift as a function of x. A LEFT~0 / RIGHT>0
gradient = cumulative justify drift (the S509 lever). A flat non-zero = a uniform x offset (margin).
A flat ~0 = horizontally aligned (gap really is weight/AA). cp932-safe: UTF-8 file, ASCII out."""
import os, sys, subprocess, io, glob
import numpy as np
from PIL import Image
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
WORD_PNG = os.path.join(ROOT, 'pipeline_data', 'word_png')
DOCX = os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx')
DPI = 150
PT = DPI / 72.0  # px per pt


def render_oxi(docx, stem, page):
    out = os.path.join('c:/tmp', 's531_' + stem[:12])
    subprocess.run([EXE, os.path.abspath(docx), out, str(DPI)], capture_output=True, text=True)
    p = out + ('_p%d.png' % page)
    return p if os.path.exists(p) else None


def load(png):
    return np.asarray(Image.open(png).convert('L'), dtype=np.float32)


def row_ink(im):
    return (255.0 - im).sum(axis=1)


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


def col_profile(im, s, e):
    # darkness per column within rows [s,e)
    return (255.0 - im[s:e, :]).sum(axis=0)


def vshift_align(w_ink, o_ink, c, half=18):
    # first remove the per-line vertical offset so the horizontal compare uses the matched rows
    lo = max(0, c - 40); hi = min(len(w_ink), c + 40)
    wseg = w_ink[lo:hi]; best = 0; bestc = -1e18
    for sh in range(-half, half + 1):
        a = lo + sh; b = hi + sh
        if a < 0 or b > len(o_ink):
            continue
        c2 = float((wseg * o_ink[a:b]).sum())
        if c2 > bestc:
            bestc = c2; best = sh
    return best


def local_hshift(wcol, ocol, cx, win=30, half=8):
    lo = max(0, cx - win); hi = min(len(wcol), cx + win)
    wseg = wcol[lo:hi]
    if wseg.sum() < 1.0:
        return None
    best = 0; bestc = -1e18
    for sh in range(-half, half + 1):
        a = lo + sh; b = hi + sh
        if a < 0 or b > len(ocol):
            continue
        c = float((wseg * ocol[a:b]).sum())
        if c > bestc:
            bestc = c; best = sh
    return best


def main():
    args = sys.argv[1:]
    page = 1
    stems = []
    for a in args:
        if a.startswith('p='):
            page = int(a[2:])
        else:
            stems.append(a)
    L = ['S531 horizontal drift in gate pixels (oxi - word). +shift = Oxi RIGHT of Word. page=%d' % page]
    for stem in stems:
        g = glob.glob(os.path.join(DOCX, stem + '*.docx'))
        wdir = glob.glob(os.path.join(WORD_PNG, stem + '*'))
        if not g or not wdir:
            L.append('%s: missing' % stem); continue
        wpng = os.path.join(wdir[0], 'page_%04d.png' % page)
        if not os.path.exists(wpng):
            L.append('%s: no word page %d' % (stem, page)); continue
        opng = render_oxi(g[0], stem, page)
        if not opng:
            L.append('%s: oxi render failed' % stem); continue
        wim = load(wpng); oim = load(opng)
        w_rink = row_ink(wim); o_rink = row_ink(oim)
        bands = line_bands(w_rink)
        L.append('--- %s : %d word text-bands ---' % (stem[:20], len(bands)))
        for bi, (s, e) in enumerate(bands):
            c = (s + e) // 2
            vsh = vshift_align(w_rink, o_rink, c)        # vertical match for this band
            os_ = max(0, s + vsh); oe_ = min(oim.shape[0], e + vsh)
            wcol = col_profile(wim, s, e)
            ocol = col_profile(oim, os_, oe_)
            # ink span of the word line
            nz = np.where(wcol > wcol.max() * 0.06)[0]
            if len(nz) < 20:
                continue
            x0, x1 = nz[0], nz[-1]
            width_pt = (x1 - x0) / PT
            # sample shift at left, mid, right thirds
            samples = []
            for frac in (0.12, 0.30, 0.50, 0.70, 0.88):
                cx = int(x0 + (x1 - x0) * frac)
                sh = local_hshift(wcol, ocol, cx)
                samples.append(sh)
            sl, _, sm, _, sr = samples
            def f(v): return ('%+d' % v) if v is not None else ' . '
            grad = (sr - sl) if (sr is not None and sl is not None) else None
            L.append('  band%2d row=%4d w=%5.1fpt vsh=%+d  L%s M%s R%s  R-L=%s px (%s pt)'
                     % (bi, c, width_pt, vsh, f(sl), f(sm), f(sr), f(grad),
                        ('%+.2f' % (grad / PT)) if grad is not None else ' . '))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s531_out.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)


if __name__ == '__main__':
    main()
