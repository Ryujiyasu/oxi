import subprocess,os,glob,json,io,numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
ROOT=os.path.abspath('.')
BIN=os.path.abspath('tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe')
base=json.load(io.open('pipeline_data/ssim_baseline.json',encoding='utf-8'))
# doc-level means, pick stratified sample
dm=sorted(((sum(p.values())/len(p),d) for d,p in base.items()))
n=len(dm)
sample=[d for _,d in dm[:9]] + [d for _,d in dm[n//2-4:n//2+4]] + [d for _,d in dm[-8:]]
def rgb(p):return np.array(Image.open(p).convert('RGB'))
def rs(o,w):
    if o.shape[:2]!=w.shape[:2]: o=np.array(Image.fromarray(o).resize((w.shape[1],w.shape[0])))
    return o
def render_score(D,ss):
    DOCX=os.path.abspath(f'tools/golden-test/documents/docx/{D}.docx')
    if not os.path.exists(DOCX): return None
    od=f'C:/tmp/net_{ss}'; os.makedirs(od,exist_ok=True)
    for f in glob.glob(od+'/*.png'): os.remove(f)
    try: subprocess.run([BIN,DOCX,od+'/oxi','150',f'--supersample={ss}'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=900)
    except: return None
    WORD=f'pipeline_data/word_png/{D}'; sc={}
    for wp in sorted(glob.glob(WORD+'/page_*.png')):
        pi=int(os.path.basename(wp)[5:9]); op=f'{od}/oxi_p{pi}.png'
        if not os.path.exists(op): continue
        w=rgb(wp); o=rs(rgb(op),w); sc[str(pi)]=ssim(w,o,channel_axis=2,data_range=255)
    return sc
tot_b=tot_n=cnt=0; rows=[]
for D in sample:
    sc2=render_score(D,2)
    if not sc2: print(f'  SKIP {D[:36]}'); continue
    bvals=base[D]
    # compare common pages
    common=[p for p in sc2 if p in bvals]
    bsum=sum(bvals[p] for p in common); nsum=sum(sc2[p] for p in common)
    tot_b+=bsum; tot_n+=nsum; cnt+=len(common)
    rows.append((nsum-bsum, D, len(common)))
print('per-doc delta (ss2 - ss1baseline), page-sum:')
for d,D,npg in sorted(rows):
    print(f'  {d:+.3f}  ({npg}pg)  {D[:40]}')
print(f'\nSAMPLE net: {cnt} pages, mean ss1={tot_b/cnt:.4f} ss2={tot_n/cnt:.4f} delta={(tot_n-tot_b)/cnt:+.4f}')
