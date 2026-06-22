# -*- coding: utf-8 -*-
"""Absolute current SSIM (DEFAULT DWrite binary) vs cached Word PNGs, no Word re-render.
Reports per-page and per-doc mean over all word_png bases. Reuses the production
_load_rgb/_resize_to_match + skimage SSIM (same as the gate)."""
import os, re, subprocess, sys, tempfile, json
from pathlib import Path
sys.path.insert(0, r"c:\Users\ryuji\oxi-main")
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim
sys.stdout.reconfigure(encoding="utf-8")
REPO=r"c:\Users\ryuji\oxi-main"
DW=os.path.join(REPO,"tools","oxi-dwrite-renderer","target","release","oxi-dwrite-renderer.exe")
DOCS=os.path.join(REPO,"tools","golden-test","documents","docx")
def ssim2(wpng,opng):
    w=_load_rgb(wpng); o=_resize_to_match(_load_rgb(opng),w)
    return ssim(w,o,full=False,channel_axis=2,data_range=255)
def find(base):
    e=Path(DOCS)/(base+".docx")
    if e.exists(): return os.path.abspath(str(e))
    c=sorted(p for p in Path(DOCS).glob(base.split("_")[0]+"*.docx") if not p.name.startswith("~$"))
    return os.path.abspath(str(c[0])) if c else None
def render(docx,outdir):
    Path(outdir).mkdir(parents=True,exist_ok=True)
    subprocess.run([DW,docx,str(Path(outdir)/"p"),str(RENDER_DPI)],capture_output=True,timeout=300)
    ps=[];i=1
    while (Path(outdir)/f"p_p{i}.png").exists(): ps.append(str(Path(outdir)/f"p_p{i}.png"));i+=1
    return ps
_bl=json.load(open(os.path.join(REPO,"pipeline_data","ssim_baseline.json"),encoding="utf-8"))
bases=sorted(_bl.keys())
page_vals=[]; doc_means=[]; missing=0; _cur={}
with tempfile.TemporaryDirectory() as tmp:
    for n,base in enumerate(bases):
        d=find(base)
        if not d: missing+=1; continue
        od=Path(tmp)/base; render(d,od)
        wdir=Path(WORD_PNG_DIR)/base; i=1; dv=[]
        while True:
            wp=wdir/f"page_{i:04d}.png"; op=od/f"p_p{i}.png"
            if not wp.exists() or not op.exists(): break
            try: dv.append(ssim2(str(wp),str(op)))
            except Exception: pass
            i+=1
        if dv:
            page_vals.extend(dv); doc_means.append(sum(dv)/len(dv)); _cur[base]={str(j+1):dv[j] for j in range(len(dv))}
        if (n+1)%40==0: print(f"  ...{n+1}/{len(bases)} done", flush=True)
page_vals.sort()
print(f"\n=== CURRENT SSIM (DWrite default binary, {len(doc_means)} docs / {len(page_vals)} pages) ===")
print(f"PER-PAGE mean = {sum(page_vals)/len(page_vals):.4f}   min = {page_vals[0]:.4f}")
print(f"PER-DOC  mean = {sum(doc_means)/len(doc_means):.4f}")
print(f"bottom-8 pages: {[round(x,3) for x in page_vals[:8]]}")
print(f"missing docx: {missing}")
json.dump(_cur,open(os.path.join(REPO,"pipeline_data","ssim_current.json"),"w"),indent=1)
print("saved pipeline_data/ssim_current.json")
