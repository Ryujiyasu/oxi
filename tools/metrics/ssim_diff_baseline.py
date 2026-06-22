# -*- coding: utf-8 -*-
"""Diff current SSIM (ssim_current.json) vs committed ssim_baseline.json, per-doc + per-page."""
import json,sys
sys.stdout.reconfigure(encoding="utf-8")
base=json.load(open("pipeline_data/ssim_baseline.json"))
cur=json.load(open("pipeline_data/ssim_current.json"))
rows=[]
for doc in sorted(set(base)|set(cur)):
    bp=base.get(doc,{}); cp=cur.get(doc,{})
    bm=sum(bp.values())/len(bp) if bp else None
    cm=sum(cp.values())/len(cp) if cp else None
    if bm is None or cm is None: 
        rows.append((doc,bm,cm,None)); continue
    rows.append((doc,bm,cm,cm-bm))
# biggest regressions / improvements
scored=[r for r in rows if r[3] is not None]
print("=== BIGGEST PER-DOC REGRESSIONS ===")
for doc,bm,cm,d in sorted(scored,key=lambda x:x[3])[:12]:
    print(f"  {doc[:40]:40} {bm:.4f} -> {cm:.4f}  Δ{d:+.4f}")
print("=== BIGGEST PER-DOC IMPROVEMENTS ===")
for doc,bm,cm,d in sorted(scored,key=lambda x:-x[3])[:12]:
    print(f"  {doc[:40]:40} {bm:.4f} -> {cm:.4f}  Δ{d:+.4f}")
# the lowest current pages
pages=[]
for doc,cp in cur.items():
    for pg,v in cp.items(): pages.append((v,doc,pg))
pages.sort()
print("=== LOWEST CURRENT PAGES ===")
for v,doc,pg in pages[:8]:
    bv=base.get(doc,{}).get(pg)
    print(f"  {doc[:38]:38} p{pg}: {v:.4f}  (baseline {bv if bv is None else round(bv,4)})")
nreg=sum(1 for r in scored if r[3]<-0.002); nimp=sum(1 for r in scored if r[3]>0.002)
print(f"\nsummary: {len(scored)} docs; {nimp} improved >0.002, {nreg} regressed <-0.002")
print(f"baseline per-doc mean {sum(r[1] for r in scored)/len(scored):.4f} -> current {sum(r[2] for r in scored)/len(scored):.4f}")
