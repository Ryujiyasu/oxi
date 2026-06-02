"""S482 gate: toggle OXI_S469_DISABLE (anchor_flow_offset natural-Y).
ON = default (S469 active), OFF = OXI_S469_DISABLE=1. Fresh DWrite vs cached Word PNG."""
import os, sys, glob, subprocess
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
REPO=r"C:\Users\ryuji\oxi-main"
DWRITE=os.path.join(REPO,"tools","oxi-dwrite-renderer","target","release","oxi-dwrite-renderer.exe")
DOCX=os.path.join(REPO,"tools","golden-test","documents","docx")
WORD=os.path.join(REPO,"pipeline_data","word_png")
TMP=r"C:\Users\ryuji\AppData\Local\Temp\s478gate"; os.makedirs(TMP,exist_ok=True)
def render(dp,pfx,on):
    env=dict(os.environ)
    if on: env.pop("OXI_S482_DISABLE",None)
    else: env["OXI_S482_DISABLE"]="1"
    subprocess.run([DWRITE,dp,pfx,"150"],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,env=env)
def sp(w,o):
    if not(os.path.exists(w) and os.path.exists(o)): return None
    a=np.array(Image.open(w).convert("L")); b=Image.open(o).convert("L")
    if b.size!=(a.shape[1],a.shape[0]): b=b.resize((a.shape[1],a.shape[0]),Image.LANCZOS)
    return float(ssim(a,np.array(b),data_range=255))
def main():
    pat=sys.argv[1] if len(sys.argv)>1 else ""
    docs=sorted(glob.glob(os.path.join(DOCX,"*%s*.docx"%pat)))
    rows=[]
    for dp in docs:
        did=os.path.splitext(os.path.basename(dp))[0]
        wdir=os.path.join(WORD,did)
        if not os.path.isdir(wdir): continue
        won=os.path.join(TMP,"on_"+did); woff=os.path.join(TMP,"off_"+did)
        render(dp,won,True); render(dp,woff,False)
        for wp in sorted(glob.glob(os.path.join(wdir,"page_*.png"))):
            pg=int(os.path.basename(wp)[5:9])
            on=sp(wp,"%s_p%d.png"%(won,pg)); off=sp(wp,"%s_p%d.png"%(woff,pg))
            if on is None or off is None: continue
            rows.append((did,pg,off,on,on-off))
    rows.sort(key=lambda r:r[4])
    for d,pg,off,on,dd in rows:
        if abs(dd)>0.0003: print("%-30s p%d OFF=%.4f ON=%.4f delta(ON-OFF)=%+.4f"%(d[:30],pg,off,on,dd))
    if rows:
        net=sum(r[4] for r in rows); up=sum(1 for r in rows if r[4]>0.0003); dn=sum(1 for r in rows if r[4]<-0.0003)
        print("-"*60); print("PAGES=%d net(ON-OFF)=%+.4f up=%d dn=%d"%(len(rows),net,up,dn))
if __name__=="__main__": main()
