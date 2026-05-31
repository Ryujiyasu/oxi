import os, sys, subprocess, glob, io, json
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
ROOT=os.path.abspath('.')
BIN=os.path.join(ROOT,'tools','oxi-dwrite-renderer','target','release','oxi-dwrite-renderer.exe')
DOCX=os.path.join(ROOT,'tools','golden-test','documents','docx','e3c545fac7a7_LOD_Handbook.docx')
DOC='e3c545fac7a7_LOD_Handbook'
WORD=os.path.join(ROOT,'pipeline_data','word_png',DOC)
def rgb(p): return np.array(Image.open(p).convert('RGB'))
def resize(o,w):
    if o.shape[:2]!=w.shape[:2]:
        o=np.array(Image.fromarray(o).resize((w.shape[1],w.shape[0])))
    return o
def render(dy, outdir):
    os.makedirs(outdir, exist_ok=True)
    for f in glob.glob(os.path.join(outdir,'*.png')): os.remove(f)
    env=dict(os.environ); env['OXI_S459_BASELINE_CJK_DY']=str(dy)
    subprocess.run([BIN, DOCX, os.path.join(outdir,'oxi'), '150'],
                   env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    # rename oxi_pN.png -> page_000N.png
    i=1
    while True:
        src=os.path.join(outdir,f'oxi_p{i}.png')
        if not os.path.exists(src): break
        os.replace(src, os.path.join(outdir,f'page_{i:04d}.png')); i+=1
def score(outdir):
    res={}
    for wp in sorted(glob.glob(os.path.join(WORD,'page_*.png'))):
        n=os.path.basename(wp)
        op=os.path.join(outdir,n)
        if not os.path.exists(op): continue
        w=rgb(wp); o=resize(rgb(op),w)
        s=ssim(w,o,channel_axis=2,data_range=255)
        res[int(n[5:9])]=s
    return res
sweeps=[float(x) for x in sys.argv[1:]] or [0.0,1.5,2.5,3.0,3.4,4.0]
allres={}
for dy in sweeps:
    od=os.path.join('C:/tmp',f's459_{dy}')
    render(dy, od); r=score(od); allres[dy]=r
    m=sum(r.values())/len(r)
    print(f'dy={dy}: mean={m:.4f}  per-page='+' '.join(f'p{k}={v:.3f}' for k,v in sorted(r.items())))
json.dump(allres, open('C:/tmp/s459_sweep.json','w'))
