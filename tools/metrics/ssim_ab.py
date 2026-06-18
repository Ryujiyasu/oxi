# -*- coding: utf-8 -*-
"""Correct, reusable SSIM A/B sentinel.

Renders every word_png base with DWrite under env A (var=1) and env B (default),
byte-compares, and SSIMs each changed doc's pages vs the cached Word reference
BOTH ways, reporting net (B - A). Uses the SAME skimage SSIM + RGB-load +
resize-to-match as the production pipeline (pipeline.ssim_calculator), so the
numbers match the gate. Earlier ad-hoc A/B scripts called calculate_ssim() with
the wrong signature (it takes dicts) and silently produced net=+0.0000 for
everything; this uses the low-level ssim directly.

Usage: python tools/metrics/ssim_ab.py OXI_S609_DISABLE   [base...]
       (A = OXI_S609_DISABLE=1, B = default. Optional base-prefix filter.)
"""
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
ARG=sys.argv[1] if len(sys.argv)>1 else "OXI_S609_DISABLE"
# "VAR" -> A sets VAR=1, B unsets. "VAR=VAL" -> A sets VAR=VAL, B unsets.
if "=" in ARG: ENVVAR,AVAL=ARG.split("=",1)
else: ENVVAR,AVAL=ARG,"1"
filt=sys.argv[2:]
def ssim2(wpng,opng):
    w=_load_rgb(wpng); o=_resize_to_match(_load_rgb(opng),w)
    return ssim(w,o,full=False,channel_axis=2,data_range=255)
bases=sorted({re.sub(r"_p\d+$","",n) for n in os.listdir(WORD_PNG_DIR)})
if filt: bases=[b for b in bases if any(b.startswith(f) for f in filt)]
def find(base):
    e=Path(DOCS)/(base+".docx")
    if e.exists(): return os.path.abspath(str(e))
    c=sorted(p for p in Path(DOCS).glob(base.split("_")[0]+"*.docx") if not p.name.startswith("~$"))
    return os.path.abspath(str(c[0])) if c else None
def render(docx,a,outdir):
    env=dict(os.environ)
    if a: env[ENVVAR]=AVAL
    else: env.pop(ENVVAR,None)
    Path(outdir).mkdir(parents=True,exist_ok=True)
    subprocess.run([DW,docx,str(Path(outdir)/"p"),str(RENDER_DPI)],capture_output=True,timeout=300,env=env)
    ps=[];i=1
    while (Path(outdir)/f"p_p{i}.png").exists(): ps.append(str(Path(outdir)/f"p_p{i}.png"));i+=1
    return ps
changed=[];checked=0
with tempfile.TemporaryDirectory() as tmp:
    seen=set()
    for base in bases:
        d=find(base)
        if not d or d in seen: continue
        seen.add(d);checked+=1
        ad=Path(tmp)/"A"/Path(d).stem; bd=Path(tmp)/"B"/Path(d).stem
        pa=render(d,True,ad); pb=render(d,False,bd)
        diff=(len(pa)!=len(pb)) or any(open(x,"rb").read()!=open(y,"rb").read() for x,y in zip(pa,pb))
        if diff: changed.append((base,ad,bd,len(pa),len(pb)))
    print(f"A={ENVVAR}={AVAL} vs B=default | checked {checked}; {len(changed)} changed bytes")
    tot=0.0; wins=losses=0; details=[]
    for base,ad,bd,na,nb in changed:
        wdir=Path(WORD_PNG_DIR)/base; i=1; net=0.0; npg=0
        while True:
            wp=wdir/f"page_{i:04d}.png"
            if not wp.exists(): break
            ap=Path(ad)/f"p_p{i}.png"; bp=Path(bd)/f"p_p{i}.png"
            if not ap.exists() or not bp.exists(): break
            try: sa=ssim2(str(wp),str(ap)); sb=ssim2(str(wp),str(bp))
            except Exception: i+=1; continue
            net+=(sb-sa); npg+=1; i+=1
        if npg==0: 
            details.append((base,None,na,nb)); continue
        tot+=net
        if net>0.0005: wins+=1
        elif net<-0.0005: losses+=1
        details.append((base,net,na,nb))
    for base,net,na,nb in sorted(details,key=lambda x:(x[1] if x[1] is not None else 0)):
        if net is None: print(f"  {base}: NO word_png pages found (pages {na}/{nb})")
        else:
            flag=" <<< REGRESS" if net<-0.0005 else (" >>> improve" if net>0.0005 else "")
            print(f"  {base}: pages={na}/{nb} net(B-A)={net:+.4f}{flag}")
    print(f"\nTOTAL net(B-A)={tot:+.4f}; improved {wins}, regressed {losses}")
