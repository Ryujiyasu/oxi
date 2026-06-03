# -*- coding: utf-8 -*-
"""S492y — map b35 p1 glyph displacement FIELD (Oxi PNG vs Word PNG) via block-matching, to
decompose the position scatter (blur-test ceiling 0.94) into SYSTEMATIC (fixable) vs RANDOM.
For each patch, find the integer (dx,dy) maximizing normalized cross-correlation in a search
window. Then fit:  dx ~ a*x (horizontal accumulation = char pitch),  dy ~ c*y (vertical
accumulation = line height),  dy ~ e*x (skew/rotation),  and report residual scatter std after
removing the linear trends. Strong slopes => fixable systematic drift. cp932-safe, ASCII out."""
import os, subprocess
from pathlib import Path
import numpy as np
from PIL import Image

R = str(Path('tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe').resolve())
DOCX = str(Path('tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx').resolve())
WD = Path('pipeline_data/word_png/b35123fe8efc_tokumei_08_01')

env = {k: v for k, v in os.environ.items() if k != 'OXI_S492X_PXSNAP'}
subprocess.run([R, DOCX, 'c:/tmp/_sf', '150'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)

PT = 150.0 / 72.0  # px per pt


def field(pg):
    wp = WD / ('page_%04d.png' % pg)
    if not wp.exists():
        return None
    w = 255.0 - np.array(Image.open(str(wp)).convert('L'), dtype=np.float32)
    sz = Image.open(str(wp)).size
    o = 255.0 - np.array(Image.open('c:/tmp/_sf_p%d.png' % pg).convert('L').resize(sz), dtype=np.float32)
    H, W = w.shape
    P, STEP, S = 28, 14, 8  # patch, step, search radius (px)
    rows = []
    for top in range(40, H - P, STEP):
        for left in range(40, W - P, STEP):
            wseg = w[top:top + P, left:left + P]
            if wseg.sum() < 800:  # skip near-empty
                continue
            best = (0, 0, -2.0)
            wm = wseg - wseg.mean()
            wn = np.sqrt((wm * wm).sum())
            if wn < 1e-3:
                continue
            for dy in range(-S, S + 1):
                for dx in range(-S, S + 1):
                    t, l = top + dy, left + dx
                    if t < 0 or l < 0 or t + P > H or l + P > W:
                        continue
                    oseg = o[t:t + P, l:l + P]
                    om = oseg - oseg.mean()
                    on = np.sqrt((om * om).sum())
                    if on < 1e-3:
                        continue
                    c = (wm * om).sum() / (wn * on)
                    if c > best[2]:
                        best = (dx, dy, c)
            if best[2] > 0.4:
                # dx,dy = how far to move OXI to match WORD at this patch
                rows.append((left + P / 2, top + P / 2, best[0], best[1], best[2]))
    return np.array(rows)


for pg in [1, 2]:
    f = field(pg)
    if f is None or len(f) < 10:
        print("p%d: insufficient patches" % pg)
        continue
    x, y, dx, dy, cc = f[:, 0], f[:, 1], f[:, 2], f[:, 3], f[:, 4]
    print("=== b35 p%d  (%d matched patches, 1pt=2.083px) ===" % (pg, len(f)))
    print("  dx: mean=%.2fpx std=%.2f   dy: mean=%.2fpx std=%.2f" % (dx.mean(), dx.std(), dy.mean(), dy.std()))

    def fit(v, t, name):
        # v ~ slope*t + b ; report slope per 100px, R^2
        A = np.vstack([t, np.ones_like(t)]).T
        sol, *_ = np.linalg.lstsq(A, v, rcond=None)
        pred = A @ sol
        ss_res = ((v - pred) ** 2).sum()
        ss_tot = ((v - v.mean()) ** 2).sum()
        r2 = 1 - ss_res / ss_tot if ss_tot > 1e-6 else 0
        return sol[0], r2
    sx, r2x = fit(dx, x, 'dx~x')
    syy, r2yy = fit(dy, y, 'dy~y')
    syx, r2yx = fit(dy, x, 'dy~x')
    sxy, r2xy = fit(dx, y, 'dx~y')
    print("  HORIZONTAL accumulation  dx~x: slope=%.4f px/100px (%.3f), R2=%.2f" % (sx * 100, sx * 100 / PT, r2x))
    print("  VERTICAL accumulation    dy~y: slope=%.4f px/100px (%.3f), R2=%.2f" % (syy * 100, syy * 100 / PT, r2yy))
    print("  SKEW (dy vs x)           dy~x: slope=%.4f px/100px, R2=%.2f" % (syx * 100, r2yx))
    print("  dx vs y                  dx~y: slope=%.4f px/100px, R2=%.2f" % (sxy * 100, r2xy))
    # residual after removing dx~x and dy~y linear trends
    Ax = np.vstack([x, np.ones_like(x)]).T
    dx_res = dx - Ax @ np.linalg.lstsq(Ax, dx, rcond=None)[0]
    Ay = np.vstack([y, np.ones_like(y)]).T
    dy_res = dy - Ay @ np.linalg.lstsq(Ay, dy, rcond=None)[0]
    print("  RESIDUAL scatter after removing linear trends: dx_res std=%.2fpx  dy_res std=%.2fpx" % (dx_res.std(), dy_res.std()))
    frac_sys = 1 - (dx_res.var() + dy_res.var()) / (dx.var() + dy.var() + 1e-6)
    print("  => systematic (linear-trend) fraction of total scatter variance: %.0f%%" % (frac_sys * 100))
    print()
