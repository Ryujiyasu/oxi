import subprocess,os,glob,sys,numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
BIN=os.path.abspath('tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe')
def rgb(p):return np.array(Image.open(p).convert('RGB'))
def rs(o,w):
    if o.shape[:2]!=w.shape[:2]: o=np.array(Image.fromarray(o).resize((w.shape[1],w.shape[0])))
    return o
def docmean(D,env):
    DOCX=os.path.abspath(f'tools/golden-test/documents/docx/{D}.docx'); WORD=f'pipeline_data/word_png/{D}'
    od='C:/tmp/cs'; os.makedirs(od,exist_ok=True)
    for f in glob.glob(od+'/*.png'): os.remove(f)
    e=dict(os.environ); e.update(env)
    subprocess.run([BIN,DOCX,od+'/oxi','150'],env=e,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=600)
    sc={}
    for wp in sorted(glob.glob(WORD+'/page_*.png')):
        pi=int(os.path.basename(wp)[5:9]); op=f'{od}/oxi_p{pi}.png'
        if os.path.exists(op):
            w=rgb(wp); o=rs(rgb(op),w); sc[pi]=ssim(w,o,channel_axis=2,data_range=255)
    return sc
docs=sys.argv[1:]
for D in docs:
    base=docmean(D,{})
    cen=docmean(D,{'OXI_S462_CELL_EXACT':'center'})
    bot=docmean(D,{'OXI_S462_CELL_EXACT':'bottom'})
    bm=sum(base.values())/len(base); cm=sum(cen.values())/len(cen); botm=sum(bot.values())/len(bot)
    print(f'{D[:30]:30s} base={bm:.4f} center={cm:.4f}({cm-bm:+.4f}) bottom={botm:.4f}({botm-bm:+.4f})')
