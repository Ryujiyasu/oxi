from PIL import Image
import numpy as np, os, json, re
from skimage.metrics import structural_similarity as ssim
def prof(a): return (a<128).sum(axis=1).astype(float)
b=json.load(open('pipeline_data/ssim_baseline.json',encoding='utf-8'))
# bottom 60 real-hex pages
allp=[(s,doc,p) for doc,pg in b.items() for p,s in pg.items() if re.match(r'^[0-9a-f]{12}_',doc)]
allp.sort()
rows=[]
for s0,doc,p in allp[:60]:
    wf=f'pipeline_data/word_png/{doc}/page_{int(p):04d}.png'; of=f'pipeline_data/oxi_png/{doc}/page_{int(p):04d}.png'
    if not(os.path.exists(wf) and os.path.exists(of)): continue
    w=np.array(Image.open(wf).convert('L'));o=np.array(Image.open(of).convert('L'))
    pw,po=prof(w),prof(o);best=(0,-1)
    for sh in range(-40,41):
        if sh>=0:a,bb=pw[sh:],po[:len(po)-sh] if sh>0 else po
        else:a,bb=pw[:len(pw)+sh],po[-sh:]
        n=min(len(a),len(bb));c=np.corrcoef(a[:n],bb[:n])[0,1]
        if c>best[1]:best=(sh,c)
    rows.append((s0,best[1],best[0],doc,p))
# sort by best-shift corr DESC among low SSIM = cheap offset candidates
print('=== CHEAP-OFFSET candidates (low SSIM, HIGH best-shift corr) ===')
cheap=[r for r in rows if r[1]>0.80 and r[0]<0.85]
for s0,bc,bs,doc,p in sorted(cheap,key=lambda x:-x[1]):
    print(f'  SSIM={s0:.3f} bestcorr={bc:.3f} shift={bs}px {doc[:34]} p{p}')
print(f'\n{len(cheap)} cheap-offset pages out of {len(rows)} bottom pages')
print('=== STRUCTURAL (low best-shift corr, hard) count:', sum(1 for r in rows if r[1]<0.6))
