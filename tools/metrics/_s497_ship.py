# -*- coding: utf-8 -*-
"""S497 ship: with the default-on binary, refresh ssim_baseline + per_glyph_baseline
for the affected doc(s), and sanity-check a control is byte-identical default vs
OXI_S497_DISABLE. cp932-safe."""
import os, json, glob, subprocess, tempfile, io, sys
import numpy as np
import fitz
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
MINCHO='C:/Windows/Fonts/msmincho.ttc';GOTHIC='C:/Windows/Fonts/msgothic.ttc';DPI=150
FF=[('Gothic',('g',GOTHIC)),('ゴシック',('g',GOTHIC)),('Goth',('g',GOTHIC)),
    ('Calibri',('cal','C:/Windows/Fonts/calibri.ttf')),('Cambria',('cam','C:/Windows/Fonts/cambria.ttc')),
    ('Times',('tim','C:/Windows/Fonts/times.ttf')),('Arial',('ari','C:/Windows/Fonts/arial.ttf')),
    ('Meiryo',('mei','C:/Windows/Fonts/meiryo.ttc')),('メイリオ',('mei','C:/Windows/Fonts/meiryo.ttc')),
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
def docx_for(did):
    for d in ['tools/golden-test/documents/docx','pipeline_data/docx']:
        for p in glob.glob(os.path.join(ROOT,d,did+'*.docx')):
            if os.path.splitext(os.path.basename(p))[0]==did:return p
        g=glob.glob(os.path.join(ROOT,d,did+'*.docx'))
        if g:return g[0]
    return None
def rgb(wp,op):
    a=np.array(Image.open(wp).convert('RGB'));b=Image.open(op).convert('RGB')
    if b.size!=(a.shape[1],a.shape[0]):b=b.resize((a.shape[1],a.shape[0]),Image.LANCZOS)
    return float(ssim(a,np.array(b),channel_axis=2,data_range=255))
def dwrite(dx,td):
    base=os.path.join(td,'d.png');subprocess.run([DW,os.path.abspath(dx),base,str(DPI)],capture_output=True,timeout=400)
    return {int(os.path.basename(f).split('_p')[-1].split('.')[0]):f for f in glob.glob(os.path.join(td,'d.png_p*.png'))}
def perglyph(dx,td):
    gj=tempfile.mktemp(suffix='.json');lj=tempfile.mktemp(suffix='.json')
    subprocess.run([DW,os.path.abspath(dx),tempfile.mktemp(),str(DPI),'--dump-glyphs='+gj],capture_output=True,timeout=400)
    subprocess.run([DW,os.path.abspath(dx),tempfile.mktemp(),str(DPI),'--dump-layout='+lj],capture_output=True,timeout=400)
    g=json.load(open(gj,encoding='utf-8'));lay=json.load(open(lj,encoding='utf-8'))
    bd={pi:[e for e in p['elements'] if e.get('type')=='border'] for pi,p in enumerate(lay['pages'])}
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
        for el in bd.get(pi,[]):
            pg.draw_line((el['x'],el['y']),(el['x']+el['w'],el['y']+el['h']),color=(0,0,0),width=0.75)
        p=os.path.join(td,f'g{pi+1}.png');pg.get_pixmap(matrix=fitz.Matrix(DPI/72,DPI/72)).save(p);out[pi+1]=p
    doc.close();return out
def measure(did,fn):
    dx=docx_for(did);wdir=os.path.join(ROOT,'pipeline_data','word_png',did)
    wpages=sorted(glob.glob(os.path.join(wdir,'page_*.png')));res={}
    with tempfile.TemporaryDirectory() as td:
        oxi=fn(dx,td)
        for wp in wpages:
            pn=int(os.path.basename(wp)[5:9])
            if pn in oxi:res[str(pn)]=rgb(wp,oxi[pn])
    return res

AFFECTED=['15076df']
# sanity: control byte-identical default vs OXI_S497_DISABLE
os.environ.pop('OXI_S497_DISABLE',None)
ctl_on=measure('b35123',dwrite)
os.environ['OXI_S497_DISABLE']='1';ctl_off=measure('b35123',dwrite);os.environ.pop('OXI_S497_DISABLE',None)
print('control b35 default vs DISABLE:', 'IDENTICAL' if ctl_on==ctl_off else ('DIFFER %s/%s'%(ctl_on,ctl_off)))

ssimb=json.load(io.open(os.path.join(ROOT,'pipeline_data','ssim_baseline.json'),encoding='utf-8'))
pgb=json.load(io.open(os.path.join(ROOT,'pipeline_data','per_glyph_baseline.json'),encoding='utf-8'))
for pref in AFFECTED:
    full=[k for k in ssimb if k.startswith(pref)][0]
    new_ss=measure(full,dwrite);new_pg=measure(full,perglyph)
    o_ss=sum(ssimb[full].values())/len(ssimb[full]);o_pg=sum(pgb[full].values())/len(pgb[full])
    ssimb[full]=new_ss;pgb[full]=new_pg
    print('%s ssim %.4f->%.4f  perglyph %.4f->%.4f'%(full[:30],o_ss,sum(new_ss.values())/len(new_ss),o_pg,sum(new_pg.values())/len(new_pg)))
json.dump(ssimb,io.open(os.path.join(ROOT,'pipeline_data','ssim_baseline.json'),'w',encoding='utf-8'),indent=2,ensure_ascii=False)
json.dump(pgb,io.open(os.path.join(ROOT,'pipeline_data','per_glyph_baseline.json'),'w',encoding='utf-8'),indent=2,ensure_ascii=False)
print('baselines patched. DONE')
