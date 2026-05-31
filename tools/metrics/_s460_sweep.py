import os, sys, subprocess, glob
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
ROOT=os.path.abspath('.')
BIN=os.path.join(ROOT,'tools','oxi-dwrite-renderer','target','release','oxi-dwrite-renderer.exe')
DOCXDIR=os.path.join(ROOT,'tools','golden-test','documents','docx')
ENVVAR=sys.argv[1]      # e.g. OXI_S455_CJK_GLYPH_DY
doc=sys.argv[2]
sweeps=[float(x) for x in sys.argv[3:]]
WORD=os.path.join(ROOT,'pipeline_data','word_png',doc)
DOCX=os.path.join(DOCXDIR,doc+'.docx')
def rgb(p): return np.array(Image.open(p).convert('RGB'))
def resize(o,w):
    if o.shape[:2]!=w.shape[:2]: o=np.array(Image.fromarray(o).resize((w.shape[1],w.shape[0])))
    return o
def run(dy):
    od=os.path.join('C:/tmp',f'sw_{doc[:8]}_{dy}')
    os.makedirs(od,exist_ok=True)
    for f in glob.glob(os.path.join(od,'*.png')): os.remove(f)
    env=dict(os.environ); env[ENVVAR]=str(dy)
    subprocess.run([BIN,DOCX,os.path.join(od,'oxi'),'150'],env=env,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    i=1
    while True:
        src=os.path.join(od,f'oxi_p{i}.png')
        if not os.path.exists(src): break
        os.replace(src,os.path.join(od,f'page_{i:04d}.png')); i+=1
    res={}
    for wp in sorted(glob.glob(os.path.join(WORD,'page_*.png'))):
        n=os.path.basename(wp); op=os.path.join(od,n)
        if not os.path.exists(op): continue
        w=rgb(wp); o=resize(rgb(op),w)
        res[int(n[5:9])]=ssim(w,o,channel_axis=2,data_range=255)
    return res
for dy in sweeps:
    r=run(dy); m=sum(r.values())/len(r)
    print(f'{ENVVAR}={dy}: mean={m:.4f}  '+' '.join(f'p{k}={v:.3f}' for k,v in sorted(r.items())))
