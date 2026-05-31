"""Direct full-corpus SSIM from cached S466 oxi_png vs word_png, vs baseline.
Bypasses the stalling recompute SSIM phase."""
import numpy as np, json, os
from PIL import Image
from skimage.metrics import structural_similarity as ssim
base=json.load(open("pipeline_data/ssim_baseline.json",encoding="utf-8"))
def sc(D,pg):
    w=f"pipeline_data/word_png/{D}/page_{pg:04d}.png"; o=f"pipeline_data/oxi_png/{D}/page_{pg:04d}.png"
    if not(os.path.exists(w) and os.path.exists(o)): return None
    a=Image.open(w).convert("L"); b=Image.open(o).convert("L")
    if b.size!=a.size: b=b.resize(a.size,Image.LANCZOS)
    return float(ssim(np.array(a),np.array(b),data_range=255))
news={}; tot=0;n=0;regs=[];imps=[]
for D in base:
    for p in base[D]:
        s=sc(D,int(p))
        if s is None: continue
        news.setdefault(D,{})[p]=s
        d=s-base[D][p]; tot+=d;n+=1
        if d<-0.003: regs.append((d,D,p))
        if d>0.003: imps.append((d,D,p))
allnew=[v for dd in news.values() for v in dd.values()]
allbase=[base[D][p] for D in news for p in news[D]]
print(f"pages scored {n}")
print(f"baseline mean {sum(allbase)/len(allbase):.4f}  S466 mean {sum(allnew)/len(allnew):.4f}  delta {(sum(allnew)-sum(allbase))/len(allnew):+.4f}")
o=sorted(allbase); nn=sorted(allnew)
for N in (3,5,10): print(f"  bottom-{N}: base {sum(o[:N]):.4f} -> S466 {sum(nn[:N]):.4f} ({sum(nn[:N])-sum(o[:N]):+.4f})")
print(f"improved>+0.003: {len(imps)}  regressed<-0.003: {len(regs)}")
print("regressions:")
for d,D,p in sorted(regs)[:15]: print(f"  {d:+.4f} {D[:40]} p{p}")
print("top improvements:")
for d,D,p in sorted(imps,reverse=True)[:10]: print(f"  {d:+.4f} {D[:40]} p{p}")
