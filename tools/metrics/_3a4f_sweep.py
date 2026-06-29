# -*- coding: utf-8 -*-
"""Per-page SSIM of 3a4f's 8 ref pages under an env override (e.g. a dy lever).
Usage: python _3a4f_sweep.py ENVVAR v1 v2 v3 ...   (each value rendered separately)
Reports per-page SSIM + 8-page mean for each value, vs the default (no override)."""
import os, sys, subprocess, tempfile
from pathlib import Path
sys.path.insert(0, r"c:\Users\ryuji\oxi-main")
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim
sys.stdout.reconfigure(encoding="utf-8")
REPO=r"c:\Users\ryuji\oxi-main"
DW=os.path.join(REPO,"tools","oxi-dwrite-renderer","target","release","oxi-dwrite-renderer.exe")
DOCX=os.path.join(REPO,"tools","golden-test","documents","docx","3a4f9fbe1a83_001620506.docx")
BASE="3a4f9fbe1a83_001620506"
ENVVAR=sys.argv[1]; VALS=sys.argv[2:]
def ssim2(wp,op):
    w=_load_rgb(wp); o=_resize_to_match(_load_rgb(op),w)
    return ssim(w,o,channel_axis=2,data_range=255)
def measure(val):
    env=dict(os.environ)
    if val=="DEFAULT": env.pop(ENVVAR,None)
    else: env[ENVVAR]=val
    with tempfile.TemporaryDirectory() as t:
        subprocess.run([DW,DOCX,str(Path(t)/"p"),str(RENDER_DPI)],capture_output=True,timeout=300,env=env)
        wdir=Path(WORD_PNG_DIR)/BASE; vals=[]
        for i in range(1,9):
            wp=wdir/f"page_{i:04d}.png"; op=Path(t)/f"p_p{i}.png"
            if wp.exists() and op.exists(): vals.append(ssim2(str(wp),str(op)))
        return vals
ref=measure("DEFAULT")
print(f"{'val':>10} " + " ".join(f"p{i+1:<5}" for i in range(len(ref))) + "  MEAN")
print(f"{'DEFAULT':>10} " + " ".join(f"{v:.3f}" for v in ref) + f"  {sum(ref)/len(ref):.4f}")
for val in VALS:
    vv=measure(val)
    deltas=" ".join(f"{vv[i]:.3f}" for i in range(len(vv)))
    print(f"{val:>10} " + deltas + f"  {sum(vv)/len(vv):.4f}   net={sum(vv)-sum(ref):+.4f}")
