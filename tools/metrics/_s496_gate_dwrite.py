# -*- coding: utf-8 -*-
"""S496 SHIP gate: dwrite-rendered PNG vs word_png RGB SSIM, ON vs OFF
(OXI_S496_TBLIND_DISABLE), for a list of docs. Mirrors the ship gate (dwrite render
incl. AA). Does NOT write any baseline. cp932-safe."""
import os, json, glob, subprocess, tempfile, sys, io
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
DPI = 150

def rgb(wp, op):
    a = np.array(Image.open(wp).convert('RGB')); b = Image.open(op).convert('RGB')
    if b.size != (a.shape[1], a.shape[0]): b = b.resize((a.shape[1], a.shape[0]), Image.LANCZOS)
    return float(ssim(a, np.array(b), channel_axis=2, data_range=255))

def docx_for(did):
    for d in ['tools/golden-test/documents/docx', 'pipeline_data/docx']:
        for p in glob.glob(os.path.join(ROOT, d, did + '*.docx')):
            if os.path.splitext(os.path.basename(p))[0] == did: return p
    g = glob.glob(os.path.join(ROOT, 'pipeline_data', '**', did + '*.docx'), recursive=True)
    return g[0] if g else None

def render(dx, td, disable):
    if disable: os.environ['OXI_S496_TBLIND_DISABLE'] = '1'
    else: os.environ.pop('OXI_S496_TBLIND_DISABLE', None)
    base = os.path.join(td, 'p.png')
    subprocess.run([DW, os.path.abspath(dx), base, str(DPI)], capture_output=True, timeout=400)
    out = {}
    for f in glob.glob(os.path.join(td, 'p.png_p*.png')):
        n = int(os.path.basename(f).split('_p')[-1].split('.')[0]); out[n] = f
    return out

def measure(dx, disable):
    wdir = os.path.join(ROOT, 'pipeline_data', 'word_png', os.path.splitext(os.path.basename(dx))[0])
    if not os.path.isdir(wdir):
        # try by did prefix
        did = os.path.splitext(os.path.basename(dx))[0]
        cand = glob.glob(os.path.join(ROOT, 'pipeline_data', 'word_png', did.split('_')[0] + '*'))
        wdir = cand[0] if cand else wdir
    wpages = sorted(glob.glob(os.path.join(wdir, 'page_*.png')))
    res = {}
    with tempfile.TemporaryDirectory() as td:
        oxi = render(dx, td, disable)
        for wp in wpages:
            pn = int(os.path.basename(wp)[5:9])
            if pn in oxi: res[pn] = rgb(wp, oxi[pn])
    return res

def main():
    ids = sys.argv[1:]
    base = json.load(io.open(os.path.join(ROOT, 'pipeline_data', 'ssim_baseline.json'), encoding='utf-8'))
    print('doc                                  page   OFF      ON     delta   (baseline)')
    tot = 0.0
    for did in ids:
        full = [k for k in base if k.startswith(did)]
        full = full[0] if full else did
        dx = docx_for(full)
        off = measure(dx, True); on = measure(dx, False)
        for pn in sorted(on):
            d = on[pn] - off.get(pn, on[pn]); tot += d
            bl = base.get(full, {}).get(str(pn), float('nan'))
            print('%-36s p%-3d %.4f  %.4f  %+.4f  (%.4f)' % (full[:36], pn, off.get(pn, float('nan')), on[pn], d, bl))
    print('TOTAL dwrite delta (ON-OFF): %+.4f' % tot)

if __name__ == '__main__':
    main()
