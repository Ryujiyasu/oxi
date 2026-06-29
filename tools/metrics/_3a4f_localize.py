# -*- coding: utf-8 -*-
"""word_png-native 2D best-shift localizer for 3a4f (the user-directed canary).

Renders 3a4f with the DEFAULT DWrite binary, then for each of its 8 reference
word_png pages computes:
  - full-page SSIM
  - per horizontal band (BAND px tall): the (dx,dy) in [-R,+R] px that maximizes
    band SSIM vs Word, and the SSIM gain. This shows WHERE the residual is and
    whether it's vertical drift (dy grows down the page), horizontal (dx),
    or structural (no shift helps).

Usage: python tools/metrics/_3a4f_localize.py [BAND] [R]
"""
import os, sys, subprocess, tempfile
from pathlib import Path
import numpy as np
sys.path.insert(0, r"c:\Users\ryuji\oxi-main")
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim
sys.stdout.reconfigure(encoding="utf-8")
REPO = r"c:\Users\ryuji\oxi-main"
DW = os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx", "3a4f9fbe1a83_001620506.docx")
BASE = "3a4f9fbe1a83_001620506"
BAND = int(sys.argv[1]) if len(sys.argv) > 1 else 80
R = int(sys.argv[2]) if len(sys.argv) > 2 else 4

def render(outdir):
    Path(outdir).mkdir(parents=True, exist_ok=True)
    subprocess.run([DW, DOCX, str(Path(outdir)/"p"), str(RENDER_DPI)], capture_output=True, timeout=300)
    ps = []; i = 1
    while (Path(outdir)/f"p_p{i}.png").exists(): ps.append(str(Path(outdir)/f"p_p{i}.png")); i += 1
    return ps

def gray(rgb):
    return (0.299*rgb[:,:,0] + 0.587*rgb[:,:,1] + 0.114*rgb[:,:,2]).astype(np.float64)

def band_ssim(w, o):
    return ssim(w, o, data_range=255)

with tempfile.TemporaryDirectory() as tmp:
    ps = render(tmp)
    wdir = Path(WORD_PNG_DIR)/BASE
    page_ssims = []
    for i in range(1, len(ps)+1):
        wp = wdir/f"page_{i:04d}.png"
        if not wp.exists(): break
        op = ps[i-1]
        W = _load_rgb(str(wp)); O = _resize_to_match(_load_rgb(op), W)
        full = ssim(W, O, full=False, channel_axis=2, data_range=255)
        page_ssims.append(full)
        Wg = gray(W); Og = gray(O)
        H, Wd = Wg.shape
        print(f"\n=== page {i}: full SSIM = {full:.4f} ===")
        print(f"  band(y0-y1) base   best(dx,dy) gain   verdict")
        dxs=[]; dys=[]
        for y0 in range(0, H, BAND):
            y1 = min(y0+BAND, H)
            wb = Wg[y0:y1, :]
            # skip near-blank bands
            if wb.std() < 3: continue
            base = band_ssim(wb, Og[y0:y1, :])
            best = base; bdx = bdy = 0
            for dy in range(-R, R+1):
                for dx in range(-R, R+1):
                    ys0 = max(0, y0+dy); ys1 = min(H, y1+dy)
                    oy0 = ys0 - (y0+dy) ; # alignment handled by slicing both
                    # shift O by (dx,dy): O_shifted[y,x] = O[y-dy, x-dx]
                    # compare wb (y0:y1) to O[y0+dy:y1+dy, dx-shift]
                    oy_a = y0+dy; oy_b = y1+dy
                    if oy_a < 0 or oy_b > H: continue
                    ob = Og[oy_a:oy_b, :]
                    if dx > 0:
                        wseg = wb[:, dx:]; oseg = ob[:, :-dx] if dx else ob
                    elif dx < 0:
                        wseg = wb[:, :dx]; oseg = ob[:, -dx:]
                    else:
                        wseg = wb; oseg = ob
                    s = band_ssim(wseg, oseg)
                    if s > best: best = s; bdx = dx; bdy = dy
            gain = best - base
            verdict = ""
            if gain > 0.02:
                if bdy != 0 and bdx == 0: verdict = "VERTICAL"
                elif bdx != 0 and bdy == 0: verdict = "HORIZONTAL"
                elif bdx != 0 and bdy != 0: verdict = "DIAG"
                dxs.append(bdx); dys.append(bdy)
            elif gain <= 0.01:
                verdict = "structural/aligned"
            print(f"  {y0:4d}-{y1:<4d} {base:.3f}  ({bdx:+d},{bdy:+d})    +{gain:.3f}  {verdict}")
        if dxs:
            import statistics
            print(f"  >> page {i} shift-needing bands: n={len(dxs)} median(dx,dy)=({int(statistics.median(dxs)):+d},{int(statistics.median(dys)):+d})  dy-range[{min(dys):+d},{max(dys):+d}]")
    print(f"\n=== 3a4f 8-page mean SSIM = {sum(page_ssims)/len(page_ssims):.4f} ===")
    print("per-page:", [round(x,4) for x in page_ssims])
