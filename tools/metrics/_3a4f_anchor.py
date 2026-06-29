# -*- coding: utf-8 -*-
"""Robust per-page anchors: first & last ink row (top anchor + bottom extent),
Word word_png vs Oxi DWrite render. 150dpi: 1px=0.48pt.
Also reports the per-page best UNIFORM vertical shift (dy maximizing full-page SSIM)."""
import os, sys, subprocess
from pathlib import Path
import numpy as np
sys.path.insert(0, r"c:\Users\ryuji\oxi-main")
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim
sys.stdout.reconfigure(encoding="utf-8")
REPO=r"c:\Users\ryuji\oxi-main"
DW=os.path.join(REPO,"tools","oxi-dwrite-renderer","target","release","oxi-dwrite-renderer.exe")
DOCX=os.path.join(REPO,"tools","golden-test","documents","docx","3a4f9fbe1a83_001620506.docx")
BASE="3a4f9fbe1a83_001620506"; OUT="C:/tmp/3a4f"; PXPT=72.0/RENDER_DPI
def gray(rgb): return (0.299*rgb[:,:,0]+0.587*rgb[:,:,1]+0.114*rgb[:,:,2]).astype(np.float64)
def first_last_ink(g):
    ink=(255.0-g).sum(axis=1); thr=ink.max()*0.06
    ys=np.where(ink>thr)[0]
    return (ys[0],ys[-1]) if len(ys) else (None,None)
Path(OUT).mkdir(parents=True,exist_ok=True)
subprocess.run([DW,DOCX,str(Path(OUT)/"p"),str(RENDER_DPI)],capture_output=True,timeout=300)
wdir=Path(WORD_PNG_DIR)/BASE
print(f"{'pg':>3} {'Wfirst':>7} {'Ofirst':>7} {'dTop':>6} | {'Wlast':>7} {'Olast':>7} {'dBot':>6} | bestUdy gain  full")
for pg in range(1,9):
    wp=wdir/f"page_{pg:04d}.png"; op=Path(OUT)/f"p_p{pg}.png"
    if not wp.exists() or not op.exists(): continue
    W=_load_rgb(str(wp)); O=_resize_to_match(_load_rgb(str(op)),W)
    wf,wl=first_last_ink(gray(W)); of,ol=first_last_ink(gray(O))
    Wg=gray(W); Og=gray(O); H=Wg.shape[0]
    full=ssim(W,O,channel_axis=2,data_range=255)
    base=ssim(Wg,Og,data_range=255); best=base; bdy=0
    for dy in range(-5,6):
        a=max(0,dy); b=min(H,H+dy)
        s=ssim(Wg[a-dy:b-dy,:] if False else Wg[max(0,-dy):H-max(0,dy) if dy>0 else H, :],
               Og[max(0,dy):H-max(0,-dy) if dy<0 else H, :], data_range=255) if False else None
        # simpler: compare W[y] to O[y+dy]
        if dy>=0: ws=Wg[:H-dy,:]; os_=Og[dy:,:]
        else: ws=Wg[-dy:,:]; os_=Og[:H+dy,:]
        s=ssim(ws,os_,data_range=255)
        if s>best: best=s; bdy=dy
    print(f"{pg:>3} {wf:>7} {of:>7} {of-wf:>+6} | {wl:>7} {ol:>7} {ol-wl:>+6} | {bdy:>+5}  {best-base:+.3f} {full:.4f}")
