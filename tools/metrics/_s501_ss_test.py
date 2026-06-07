# -*- coding: utf-8 -*-
"""S501 weight/AA: test supersample 2 (current default) vs 3 vs 4 on the weight-capped
bottom-N docs. Render each via dwrite --supersample=N, SSIM vs word_png. cp932-safe."""
import os, sys, glob, subprocess, tempfile, io
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')

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

def render(dx, ss, td):
    base = os.path.join(td, f's{ss}.png')
    subprocess.run([DW, os.path.abspath(dx), base, '150', f'--supersample={ss}'], capture_output=True, timeout=600)
    return {int(os.path.basename(f).split('_p')[-1].split('.')[0]): f for f in glob.glob(os.path.join(td, f's{ss}.png_p*.png'))}

def docmean(dx, did, ss):
    wdir = os.path.join(ROOT, 'pipeline_data', 'word_png', did)
    wpages = sorted(glob.glob(os.path.join(wdir, 'page_*.png')))
    with tempfile.TemporaryDirectory() as td:
        oxi = render(dx, ss, td)
        vals = []
        for wp in wpages:
            pn = int(os.path.basename(wp)[5:9])
            if pn in oxi: vals.append(rgb(wp, oxi[pn]))
    return sum(vals) / len(vals) if vals else 0.0

def main():
    ids = sys.argv[1:]
    print('doc                          ss2       ss3       ss4    (d3-2, d4-2)')
    t2 = t3 = t4 = 0.0
    for did in ids:
        dx = docx_for(did)
        full = os.path.splitext(os.path.basename(dx))[0]
        m2 = docmean(dx, full, 2); m3 = docmean(dx, full, 3); m4 = docmean(dx, full, 4)
        t2 += m2; t3 += m3; t4 += m4
        print('%-26s %.4f  %.4f  %.4f   (%+.4f, %+.4f)' % (full[:26], m2, m3, m4, m3 - m2, m4 - m2))
    print('TOTAL                      %.4f  %.4f  %.4f   (%+.4f, %+.4f)' % (t2, t3, t4, t3 - t2, t4 - t2))

if __name__ == '__main__':
    main()
