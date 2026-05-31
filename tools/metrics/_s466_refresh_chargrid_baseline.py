"""Refresh baseline for charGrid (linesAndChars) docs ONLY, to the current
default-ON renders' absolute SSIM. Non-charGrid pages are left untouched (S466
does not affect them; their staleness is a separate ratchet-only-baseline
issue). Renders each charGrid doc fresh with default-ON DWrite."""
import json, os, zipfile, glob, subprocess, tempfile
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
ROOT=os.path.abspath(os.path.join(os.path.dirname(__file__),"..",".."))
DW=os.path.join(ROOT,"tools","oxi-dwrite-renderer","target","release","oxi-dwrite-renderer.exe")
DOCX=os.path.join(ROOT,"tools","golden-test","documents","docx")
WPNG=os.path.join(ROOT,"pipeline_data","word_png")
BASE=os.path.join(ROOT,"pipeline_data","ssim_baseline.json")
base=json.load(open(BASE,encoding="utf-8"))
def is_chargrid(D):
    fs=glob.glob(os.path.join(DOCX,D+".docx"))
    if not fs: return False
    try:
        z=zipfile.ZipFile(fs[0]); return 'w:type="linesAndChars"' in z.read("word/document.xml").decode("utf-8")
    except: return False
def score(wp,op):
    if not(os.path.exists(wp) and os.path.exists(op)): return None
    a=Image.open(wp).convert("L"); b=Image.open(op).convert("L")
    if b.size!=a.size: b=b.resize(a.size,Image.LANCZOS)
    return float(ssim(np.array(a),np.array(b),data_range=255))
cg=[D for D in base if is_chargrid(D)]
print(f"refreshing {len(cg)} charGrid docs")
changes=[]; skipped=0
for D in cg:
    docx=os.path.join(DOCX,D+".docx")
    if not os.path.exists(os.path.join(WPNG,D)): skipped+=1; continue  # no Word ref
    with tempfile.TemporaryDirectory() as td:
        pref=os.path.join(td,"oxi")
        subprocess.run([DW,docx,pref,"150"],capture_output=True,timeout=120)
        for pg in list(base[D].keys()):
            wp=os.path.join(WPNG,D,f"page_{int(pg):04d}.png")
            op=f"{pref}_p{int(pg)}.png"
            s=score(wp,op)
            if s is None: continue
            old=base[D][pg]
            if abs(s-old)>1e-6:
                changes.append((D,pg,old,s)); base[D][pg]=s
json.dump(base,open(BASE,"w",encoding="utf-8"),indent=2,ensure_ascii=False)
a=[v for p in base.values() for v in p.values()]
print(f"updated {len(changes)} pages (skipped {skipped} no-word-ref). new baseline mean {sum(a)/len(a):.4f}")
for D,pg,o,n in sorted(changes,key=lambda x:x[3]-x[2])[:8]: print(f"  DOWN {D[:34]} p{pg} {o:.4f}->{n:.4f}")
for D,pg,o,n in sorted(changes,key=lambda x:-(x[3]-x[2]))[:8]: print(f"  UP   {D[:34]} p{pg} {o:.4f}->{n:.4f}")
