import subprocess, os, numpy as np, glob, json, re
from PIL import Image
from skimage.metrics import structural_similarity as ssim
DW=os.path.abspath('tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe')
TMP=os.path.abspath('C:/tmp/s454g'); os.makedirs(TMP,exist_ok=True)
b=json.load(open('pipeline_data/ssim_baseline.json',encoding='utf-8'))
real=[d for d in b if re.match(r'^[0-9a-f]{12}_',d)]
def render(D,delta,tag):
    src=os.path.abspath(f'tools/golden-test/documents/docx/{D}.docx')
    out=os.path.join(TMP,tag); env=dict(os.environ); env['OXI_S455_CJK_GLYPH_DY']=str(delta)
    subprocess.run([DW,src,out,'150'],env=env,capture_output=True,text=True); return out
def load(f): return np.array(Image.open(f).convert('L'))
def docssim(out,D):
    npg=len(glob.glob(f'pipeline_data/word_png/{D}/page_*.png')); per=[]
    for pg in range(1,npg+1):
        of=f'{out}_p{pg}.png'; wf=f'pipeline_data/word_png/{D}/page_{pg:04d}.png'
        if not(os.path.exists(of) and os.path.exists(wf)): continue
        o=load(of);w=load(wf);H=min(o.shape[0],w.shape[0]);W=min(o.shape[1],w.shape[1])
        per.append(ssim(w[:H,:W],o[:H,:W]))
    for pg in range(1,30):
        f=f'{out}_p{pg}.png'
        if os.path.exists(f): os.remove(f)
    return per
DELTA=1.5
rows=[]
allp0=[]; allp1=[]
for D in real:
    if not os.path.exists(f'tools/golden-test/documents/docx/{D}.docx'):
        continue
    p0=docssim(render(D,0.0,'a'),D)
    p1=docssim(render(D,DELTA,'b'),D)
    if not p0: continue
    allp0+=p0; allp1+=p1
    d=np.mean(p1)-np.mean(p0)
    rows.append((D,np.mean(p0),np.mean(p1),d))
rows.sort(key=lambda r:r[3])
print(f'=== S454 LM0 glyph dy={DELTA} : full real corpus (DWrite) ===')
print(f'{"doc":30} {"d0":>6} {"d1.5":>6} {"delta":>7}')
for D,m0,m1,d in rows:
    flag='  <== REGRESS' if d<-0.001 else ('  <== improve' if d>0.001 else '')
    print(f'{D[:30]:30} {m0:6.3f} {m1:6.3f} {d:+7.4f}{flag}')
print(f'\nPER-PAGE mean: d0={np.mean(allp0):.4f} d{DELTA}={np.mean(allp1):.4f} ({np.mean(allp1)-np.mean(allp0):+.4f})')
b5=sorted(allp0)[:5]; b5n=sorted(allp1)[:5]
print(f'bottom-5 page sum: d0={sum(b5):.4f} d{DELTA}={sum(b5n):.4f} ({sum(b5n)-sum(b5):+.4f})')
nreg=sum(1 for r in rows if r[3]<-0.001); nimp=sum(1 for r in rows if r[3]>0.001)
print(f'docs: {nimp} improved, {nreg} regressed, {len(rows)-nimp-nreg} flat')
