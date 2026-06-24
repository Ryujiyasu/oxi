import os, sys, io, json, zipfile, subprocess
import numpy as np
from PIL import Image
sys.path.insert(0,'tools/metrics')
from mixedh_lineplace import MATH, build_math
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="backslashreplace")
ROOT=os.path.abspath('.')
EXE=os.path.join(ROOT,'tools','oxi-dwrite-renderer','target','release','oxi-dwrite-renderer.exe')
DPI=300
# Word T->B targets (pixel, from earlier measurement)
WTARGET={'x':24.48,'sup':24.72,'frac':33.84,'nest':42.0,'rad':26.16,'sum':46.56,'matrix':33.36}
def ink_bands(png,dpi):
    im=np.asarray(Image.open(png).convert('L'),dtype=np.float32)
    rows=(im<128).sum(axis=1)>0
    bands=[]; i=0; n=len(rows)
    while i<n:
        if rows[i]:
            j=i
            while j<n and (rows[j] or (j+1<n and rows[j+1])): j+=1
            bands.append((i,j-1)); i=j
        else: i+=1
    return [(t/dpi*72.0,b/dpi*72.0) for t,b in bands]
def otb(name, env):
    dx=build_math('sw_'+name+'.docx', MATH[name])
    op='c:/tmp/sw_'+name
    e=dict(os.environ); e.update(env)
    subprocess.run([EXE,os.path.abspath(dx),op,str(DPI)],capture_output=True,text=True,env=e)
    b=ink_bands(op+'_p1.png',DPI)
    if len(b)<3: return None
    return round(b[-1][0]-b[0][0],2)
combos=[(asc,lead) for asc in (0.55,0.60,0.65) for lead in (0.0,1.0,1.5,2.0)]
print('%-12s | %s'%('combo(asc,lead)', '  '.join('%-6s'%k for k in WTARGET)+' | maxabs rms'))
for asc,lead in combos:
    env={'OXI_S529_ASC':str(asc),'OXI_S529_LEAD':str(lead),'OXI_S529_DESC':'0.05','OXI_S529_FLOOR':'1.14'}
    errs=[]; cells=[]
    for k in WTARGET:
        o=otb(k,env)
        d=round(o-WTARGET[k],2) if o is not None else None
        errs.append(d); cells.append('%+5.2f'%d if d is not None else ' n/a ')
    valid=[e for e in errs if e is not None]
    mx=max(abs(e) for e in valid); rms=(sum(e*e for e in valid)/len(valid))**0.5
    print('asc%.2f l%.1f | %s | %.2f %.2f'%(asc,lead,'  '.join(cells),mx,rms))
