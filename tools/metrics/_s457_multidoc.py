# -*- coding: utf-8 -*-
"""Sweep an env var across several word_png docs, report per-doc mean SSIM.
Usage: python _s457_multidoc.py ENVVAR v1 v2 ... :: base1 base2 ..."""
import os, sys, subprocess, tempfile
from pathlib import Path
sys.path.insert(0, r"c:\Users\ryuji\oxi-main")
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim
sys.stdout.reconfigure(encoding="utf-8")
REPO=r"c:\Users\ryuji\oxi-main"
DW=os.path.join(REPO,"tools","oxi-dwrite-renderer","target","release","oxi-dwrite-renderer.exe")
DOCS=os.path.join(REPO,"tools","golden-test","documents","docx")
args=sys.argv[1:]; sep=args.index("::"); ENVVAR=args[0]; VALS=args[1:sep]; BASES=args[sep+1:]
def docpath(base):
    p=Path(DOCS)/(base+".docx")
    if p.exists(): return str(p)
    c=sorted(x for x in Path(DOCS).glob(base.split("_")[0]+"*.docx") if not x.name.startswith("~$"))
    return str(c[0]) if c else None
def ssim2(wp,op):
    w=_load_rgb(wp); o=_resize_to_match(_load_rgb(op),w); return ssim(w,o,channel_axis=2,data_range=255)
def measure(base,val):
    d=docpath(base); env=dict(os.environ)
    if val=="DEFAULT": env.pop(ENVVAR,None)
    else: env[ENVVAR]=val
    with tempfile.TemporaryDirectory() as t:
        subprocess.run([DW,d,str(Path(t)/"p"),str(RENDER_DPI)],capture_output=True,timeout=400,env=env)
        wdir=Path(WORD_PNG_DIR)/base; vals=[]; i=1
        while True:
            wp=wdir/f"page_{i:04d}.png"; op=Path(t)/f"p_p{i}.png"
            if not wp.exists() or not op.exists(): break
            vals.append(ssim2(str(wp),str(op))); i+=1
        return sum(vals)/len(vals) if vals else None
print(f"ENV {ENVVAR}; vals {['DEFAULT']+VALS}")
print(f"{'doc':<28} " + " ".join(f"{v:>8}" for v in ['DEFAULT']+VALS))
for base in BASES:
    row=[measure(base,v) for v in ['DEFAULT']+VALS]
    print(f"{base[:28]:<28} " + " ".join(f"{(x if x is not None else 0):.4f}  " for x in row))
