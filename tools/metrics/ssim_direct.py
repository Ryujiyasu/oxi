# -*- coding: utf-8 -*-
import os, re, subprocess, sys, tempfile
from pathlib import Path
sys.path.insert(0, r"c:\Users\ryuji\oxi-main")
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim
sys.stdout.reconfigure(encoding="utf-8")
REPO=r"c:\Users\ryuji\oxi-main"
DW=os.path.join(REPO,"tools","oxi-dwrite-renderer","target","release","oxi-dwrite-renderer.exe")
DOCS=os.path.join(REPO,"tools","golden-test","documents","docx")
spec=sys.argv[1]; filt=sys.argv[2:]
env_a={}
if "=" in spec: k,v=spec.split("=",1); env_a[k]=v
else: env_a[spec]="1"
def ssim2(w,o):
    W=_load_rgb(w); O=_resize_to_match(_load_rgb(o),W)
    return ssim(W,O,full=False,channel_axis=2,data_range=255)
def find(base):
    e=Path(DOCS)/(base+".docx")
    if e.exists(): return os.path.abspath(str(e))
    c=sorted(p for p in Path(DOCS).glob(base.split("_")[0]+"*.docx") if not p.name.startswith("~$"))
    return os.path.abspath(str(c[0])) if c else None
bases=sorted({re.sub(r"_p\d+$","",n) for n in os.listdir(WORD_PNG_DIR)})
if filt: bases=[b for b in bases if any(f in b for f in filt)]
def render(d,setflag,outdir):
    env=dict(os.environ)
    if setflag:
        for k,v in env_a.items(): env[k]=v
    else:
        for k in env_a: env.pop(k,None)
    Path(outdir).mkdir(parents=True,exist_ok=True)
    subprocess.run([DW,d,str(Path(outdir)/"p"),str(RENDER_DPI)],capture_output=True,timeout=300,env=env)
    return outdir
TOT=0.0; up=dn=0
print(f"spec={spec}  bases={len(bases)}")
with tempfile.TemporaryDirectory() as tmp:
    for base in bases:
        d=find(base)
        if not d: continue
        wdir=Path(WORD_PNG_DIR)/base
        if not (wdir/"page_0001.png").exists(): continue
        on=render(d,True,Path(tmp)/"on"/base); off=render(d,False,Path(tmp)/"off"/base)
        i=1; net=0.0; n=0
        while (wdir/f"page_{i:04d}.png").exists():
            ap=Path(on)/f"p_p{i}.png"; bp=Path(off)/f"p_p{i}.png"
            if not ap.exists() or not bp.exists(): break
            try: a=ssim2(str(wdir/f'page_{i:04d}.png'),str(ap)); b=ssim2(str(wdir/f'page_{i:04d}.png'),str(bp))
            except Exception: i+=1; continue
            net+=(a-b); n+=1; i+=1
        if n==0: continue
        TOT+=net
        if net>0.0005: up+=1; tag=">>> IMPROVE"
        elif net<-0.0005: dn+=1; tag="<<< REGRESS"
        else: tag=""
        if abs(net)>0.0003: print(f"  {base}: net ON-OFF={net:+.4f} ({n}pg) {tag}")
print(f"\nTOTAL net ON-OFF={TOT:+.4f}; improved {up}, regressed {dn}  (positive=flag IMPROVES)")
