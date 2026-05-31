"""Direct SSIM for the charGrid family (+controls) from CACHED S466 oxi_png vs
word_png, compared to baseline. Bypasses the flaky recompute pipeline."""
import json, glob, os
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
ROOT="pipeline_data"
base=json.load(open(f"{ROOT}/ssim_baseline.json",encoding="utf-8"))
def score(doc,pg):
    w=f"{ROOT}/word_png/{doc}/page_{int(pg):04d}.png"
    o=f"{ROOT}/oxi_png/{doc}/page_{int(pg):04d}.png"
    if not (os.path.exists(w) and os.path.exists(o)): return None
    a=Image.open(w).convert("L"); b=Image.open(o).convert("L")
    if b.size!=a.size: b=b.resize(a.size,Image.LANCZOS)
    return float(ssim(np.array(a),np.array(b),data_range=255))
# charGrid family doc_ids (substring match)
families={"tokumei":"tokumei","b35":"b35123","b837":"b837808","1636":"1636d28","15076":"15076df","6514":"6514f2","de6e":"de6e32","a1d6":"a1d6e4","29dc":"29dc6e"}
print(f"{'doc/page':<48}{'base':>8}{'S466':>8}{'delta':>8}")
print("-"*74)
tot_d=0; n=0; chg=[]
for doc in sorted(base):
    if not any(s in doc for s in families.values()): continue
    for pg in sorted(base[doc],key=lambda x:int(x)):
        b0=base[doc][pg]; s=score(doc,pg)
        if s is None: continue
        d=s-b0; tot_d+=d; n+=1; chg.append((d,doc,pg,b0,s))
        if abs(d)>0.003:
            print(f"{doc[:40]+' p'+pg:<48}{b0:>8.4f}{s:>8.4f}{d:>+8.4f}")
print("-"*74)
print(f"charGrid family: {n} pages, mean delta {tot_d/n:+.4f}" if n else "no pages")
up=sum(1 for d,*_ in chg if d>0.003); dn=sum(1 for d,*_ in chg if d<-0.003)
print(f"improved {up}, regressed {dn}, ~same {n-up-dn}")
