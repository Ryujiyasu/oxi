# -*- coding: utf-8 -*-
"""S497 gate: per-glyph (AA-free) AND dwrite SSIM, with OXI_S499_EST_RENDER_LH ON vs OFF,
for a list of docs. S497 default is OFF, so OFF == current baseline. Does NOT write
baselines. cp932-safe."""
import os, json, glob, subprocess, tempfile, sys, io
import numpy as np
import fitz
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
MINCHO = 'C:/Windows/Fonts/msmincho.ttc'; GOTHIC = 'C:/Windows/Fonts/msgothic.ttc'
DPI = 150
FF = [('Gothic',('g',GOTHIC)),('繧ｴ繧ｷ繝・け',('g',GOTHIC)),('Goth',('g',GOTHIC)),
    ('Calibri',('cal','C:/Windows/Fonts/calibri.ttf')),('Cambria',('cam','C:/Windows/Fonts/cambria.ttc')),
    ('Times',('tim','C:/Windows/Fonts/times.ttf')),('Arial',('ari','C:/Windows/Fonts/arial.ttf')),
    ('Meiryo',('mei','C:/Windows/Fonts/meiryo.ttc')),('繝｡繧､繝ｪ繧ｪ',('mei','C:/Windows/Fonts/meiryo.ttc')),
    ('Yu Gothic',('yg','C:/Windows/Fonts/YuGothM.ttc')),('Yu Mincho',('ym','C:/Windows/Fonts/yumin.ttf'))]
_FC={}
def _ff(nm,p):
    if nm not in _FC:
        try:_FC[nm]=fitz.Font(fontfile=p)
        except Exception:_FC[nm]=fitz.Font(fontfile=MINCHO)
    return _FC[nm]
K=0.859
def _fp(fam):
    fam=fam or ''
    for n,(nm,p) in FF:
        if n in fam:return nm,p
    return 'm',MINCHO

def perglyph_pngs(docx,outdir):
    gj=tempfile.mktemp(suffix='.json');lj=tempfile.mktemp(suffix='.json')
    subprocess.run([DW,os.path.abspath(docx),tempfile.mktemp(),str(DPI),'--dump-glyphs='+gj],capture_output=True,timeout=400)
    subprocess.run([DW,os.path.abspath(docx),tempfile.mktemp(),str(DPI),'--dump-layout='+lj],capture_output=True,timeout=400)
    g=json.load(open(gj,encoding='utf-8'));lay=json.load(open(lj,encoding='utf-8'))
    for f in (gj,lj):
        if os.path.exists(f):os.unlink(f)
    borders={pi:[e for e in p['elements'] if e.get('type')=='border'] for pi,p in enumerate(lay['pages'])}
    out={};doc=fitz.open()
    for pi,page in enumerate(g['pages']):
        pg=doc.new_page(width=page['width'],height=page['height']);loaded=set()
        for gl in page['glyphs']:
            fn,path=_fp(gl.get('font_family'));fs=gl['font_size']
            if fn not in loaded:
                try:pg.insert_font(fontname=fn,fontfile=path);loaded.add(fn)
                except Exception:pass
            by=gl.get('baseline',gl['top']+fs*K)
            try:pg.insert_text((gl['x'],by),gl['char'],fontname=fn,fontsize=fs,color=(0,0,0))
            except Exception:pass
        for el in borders.get(pi,[]):
            pg.draw_line((el['x'],el['y']),(el['x']+el['w'],el['y']+el['h']),color=(0,0,0),width=0.75)
        p=os.path.join(outdir,f'g{pi+1}.png');pg.get_pixmap(matrix=fitz.Matrix(DPI/72,DPI/72)).save(p);out[pi+1]=p
    doc.close();return out

def dwrite_pngs(docx,outdir):
    base=os.path.join(outdir,'d.png')
    subprocess.run([DW,os.path.abspath(docx),base,str(DPI)],capture_output=True,timeout=400)
    out={}
    for f in glob.glob(os.path.join(outdir,'d.png_p*.png')):
        n=int(os.path.basename(f).split('_p')[-1].split('.')[0]);out[n]=f
    return out

def rgb(wp,op):
    a=np.array(Image.open(wp).convert('RGB'));b=Image.open(op).convert('RGB')
    if b.size!=(a.shape[1],a.shape[0]):b=b.resize((a.shape[1],a.shape[0]),Image.LANCZOS)
    return float(ssim(a,np.array(b),channel_axis=2,data_range=255))

def docx_for(did):
    for d in ['tools/golden-test/documents/docx','pipeline_data/docx']:
        for p in glob.glob(os.path.join(ROOT,d,did+'*.docx')):
            if os.path.splitext(os.path.basename(p))[0]==did:return p
    g=glob.glob(os.path.join(ROOT,'pipeline_data','**',did+'*.docx'),recursive=True)
    return g[0] if g else None

def measure(did,on,mode):
    if on:os.environ['OXI_S499_EST_RENDER_LH']='1'
    else:os.environ.pop('OXI_S499_EST_RENDER_LH',None)
    dx=docx_for(did)
    wdir=os.path.join(ROOT,'pipeline_data','word_png',did)
    wpages=sorted(glob.glob(os.path.join(wdir,'page_*.png')))
    res={}
    with tempfile.TemporaryDirectory() as td:
        oxi=perglyph_pngs(dx,td) if mode=='pg' else dwrite_pngs(dx,td)
        for wp in wpages:
            pn=int(os.path.basename(wp)[5:9])
            if pn in oxi:res[pn]=rgb(wp,oxi[pn])
    return res

def main():
    mode=sys.argv[1]  # 'pg' or 'dw'
    ids=sys.argv[2:]
    bfile='per_glyph_baseline.json' if mode=='pg' else 'ssim_baseline.json'
    base=json.load(io.open(os.path.join(ROOT,'pipeline_data',bfile),encoding='utf-8'))
    print('mode=%s  doc                            page  OFF      ON     delta'%mode)
    tot=0.0
    for did in ids:
        full=[k for k in base if k.startswith(did)];full=full[0] if full else did
        off=measure(full,False,mode);on=measure(full,True,mode)
        dm_off=sum(off.values())/len(off) if off else 0
        dm_on=sum(on.values())/len(on) if on else 0
        for pn in sorted(on):
            d=on[pn]-off.get(pn,on[pn]);tot+=d
            print('%-32s p%-3d %.4f  %.4f  %+.4f'%(full[:32],pn,off.get(pn,float('nan')),on[pn],d))
        print('   -> %s docmean %.4f -> %.4f (%+.4f)'%(full[:24],dm_off,dm_on,dm_on-dm_off))
    print('TOTAL %s delta (ON-OFF): %+.4f'%(mode,tot))

if __name__=='__main__':
    main()

