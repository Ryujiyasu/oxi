# -*- coding: utf-8 -*-
"""S501 ss3 gate: render given docs at --supersample=3, compare doc-mean to the ss2
ssim_baseline.json. Reports bottom-N sum delta + per-doc. Non-destructive. cp932-safe."""
import os, sys, glob, subprocess, tempfile, io, json
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
BASE = json.load(io.open(os.path.join(ROOT, 'pipeline_data', 'ssim_baseline.json'), encoding='utf-8'))

def docx_for(did):
    for d in ['tools/golden-test/documents/docx', 'pipeline_data/docx']:
        for p in glob.glob(os.path.join(ROOT, d, did + '*.docx')):
            if os.path.splitext(os.path.basename(p))[0] == did: return p
        g = glob.glob(os.path.join(ROOT, d, did + '*.docx'))
        if g: return g[0]
    g = glob.glob(os.path.join(ROOT, 'pipeline_data', '**', did + '*.docx'), recursive=True)
    return g[0] if g else None

def rgb(wp, op):
    a = np.array(Image.open(wp).convert('RGB')); b = Image.open(op).convert('RGB')
    if b.size != (a.shape[1], a.shape[0]): b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), channel_axis=2, data_range=255))

def ss3_mean(dx, did):
    wdir = os.path.join(ROOT, 'pipeline_data', 'word_png', did)
    wp = sorted(glob.glob(os.path.join(wdir, 'page_*.png')))
    with tempfile.TemporaryDirectory() as td:
        base = os.path.join(td, 'p.png')
        subprocess.run([DW, os.path.abspath(dx), base, '150', '--supersample=3'], capture_output=True, timeout=600)
        oxi = {int(os.path.basename(f).split('_p')[-1].split('.')[0]): f for f in glob.glob(os.path.join(td, 'p.png_p*.png'))}
        v = [rgb(w, oxi[int(os.path.basename(w)[5:9])]) for w in wp if int(os.path.basename(w)[5:9]) in oxi]
    return sum(v) / len(v) if v else None

def main():
    # default: the full ssim bottom-20
    if len(sys.argv) > 1:
        prefs = sys.argv[1:]
    else:
        docs = sorted(((sum(v.values()) / len(v), k) for k, v in BASE.items()))[:20]
        prefs = [k for _, k in docs]
    tot2 = tot3 = 0.0; rows = []
    for pref in prefs:
        full = [k for k in BASE if k.startswith(pref)]
        full = full[0] if full else pref
        dx = docx_for(full)
        if not dx: continue
        m2 = sum(BASE[full].values()) / len(BASE[full])
        m3 = ss3_mean(dx, full)
        if m3 is None: continue
        tot2 += m2; tot3 += m3; rows.append((full, m2, m3))
    rows.sort(key=lambda r: r[1])
    print('doc                          ss2(base)  ss3      delta')
    for full, m2, m3 in rows:
        print('%-26s %.4f    %.4f  %+.4f' % (full[:26], m2, m3, m3 - m2))
    print('BOTTOM-%d SUM  ss2=%.4f  ss3=%.4f  delta=%+.4f  (%d up / %d down)' %
          (len(rows), tot2, tot3, tot3 - tot2,
           sum(1 for _, a, b in rows if b > a + 1e-5), sum(1 for _, a, b in rows if b < a - 1e-5)))

if __name__ == '__main__':
    main()
