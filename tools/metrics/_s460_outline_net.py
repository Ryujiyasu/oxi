import subprocess,os,glob,json,io,numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
BIN=os.path.abspath('tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe')
base=json.load(io.open('pipeline_data/ssim_baseline.json',encoding='utf-8'))
dm=sorted(((sum(p.values())/len(p),d) for d,p in base.items()))
n=len(dm)
# wider sample: 12 bottom, 10 mid, 12 top
sample=[d for _,d in dm[:12]]+[d for _,d in dm[n//2-5:n//2+5]]+[d for _,d in dm[-12:]]
def rgb(p):return np.array(Image.open(p).convert('RGB'))
def rs(o,w):
    if o.shape[:2]!=w.shape[:2]: o=np.array(Image.fromarray(o).resize((w.shape[1],w.shape[0])))
    return o
def score(D):
    DOCX=os.path.abspath(f'tools/golden-test/documents/docx/{D}.docx')
    if not os.path.exists(DOCX): return None
    od='C:/tmp/oln'; os.makedirs(od,exist_ok=True)
    for f in glob.glob(od+'/*.png'): os.remove(f)
    e=dict(os.environ); e['OXI_S460_RMODE']='outline'
    try: subprocess.run([BIN,DOCX,od+'/oxi','150'],env=e,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=600)
    except: return None
    WORD=f'pipeline_data/word_png/{D}'; sc={}
    for wp in sorted(glob.glob(WORD+'/page_*.png')):
        pi=str(int(os.path.basename(wp)[5:9])); op=f'{od}/oxi_p{pi}.png'
        if not os.path.exists(op): continue
        w=rgb(wp); o=rs(rgb(op),w); sc[pi]=ssim(w,o,channel_axis=2,data_range=255)
    return sc
tb=tn=cnt=0; rows=[]
for D in sample:
    sc=score(D)
    if not sc: continue
    bv=base[D]; common=[p for p in sc if p in bv]
    bs=sum(bv[p] for p in common); ns=sum(sc[p] for p in common)
    tb+=bs; tn+=ns; cnt+=len(common); rows.append(((ns-bs)/len(common),D,len(common)))
print('per-doc mean delta (outline_ss1 - ss1baseline):')
for d,D,npg in sorted(rows):
    print(f'  {d:+.4f}  ({npg}pg)  {D[:40]}')
print(f'\nSAMPLE net: {cnt}pg mean ss1={tb/cnt:.4f} outline={tn/cnt:.4f} delta={(tn-tb)/cnt:+.4f}')
print(f'regressions (<-0.002): {sum(1 for d,_,_ in rows if d<-0.002)}  improvements (>+0.002): {sum(1 for d,_,_ in rows if d>0.002)}')
