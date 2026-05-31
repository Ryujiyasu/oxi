import subprocess,os,glob,sys,json,numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
ROOT=os.path.abspath('.')
BIN=os.path.abspath('tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe')
def rgb(p):return np.array(Image.open(p).convert('RGB'))
def rs(o,w):
    if o.shape[:2]!=w.shape[:2]: o=np.array(Image.fromarray(o).resize((w.shape[1],w.shape[0])))
    return o
def docssim(D,ss):
    WORD=f'pipeline_data/word_png/{D}'
    DOCX=os.path.abspath(f'tools/golden-test/documents/docx/{D}.docx')
    od=f'C:/tmp/ms_{ss}'; os.makedirs(od,exist_ok=True)
    for f in glob.glob(od+'/*.png'): os.remove(f)
    subprocess.run([BIN,DOCX,od+'/oxi','150',f'--supersample={ss}'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=600)
    sc={}
    for wp in sorted(glob.glob(WORD+'/page_*.png')):
        pi=int(os.path.basename(wp)[5:9]); op=f'{od}/oxi_p{pi}.png'
        if not os.path.exists(op): continue
        w=rgb(wp); o=rs(rgb(op),w); sc[pi]=ssim(w,o,channel_axis=2,data_range=255)
    return sum(sc.values())/len(sc) if sc else 0
docs=sys.argv[1:]
for D in docs:
    r={ss:docssim(D,ss) for ss in [1,2]}
    print(f'{D[:34]:34s} ss1={r[1]:.4f} ss2={r[2]:.4f} delta={r[2]-r[1]:+.4f}')
