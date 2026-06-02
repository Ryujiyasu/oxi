"""S492 SSIM sentinel — OFF vs ON on the docs F1 actually touches (type=lines with
jc=left wrapping content). Renders each per-page slice with the DEFAULT renderer
(DWrite) under OXI_S492_JCNATURAL unset vs =1, computes full-page SSIM vs the cached
Word ground-truth PNG (pipeline_data/per_page_ssim/word_png/<stem>_p1.png), reports
per-page OFF/ON/delta + net. Run after a Word-PNG cache exists (per_page_ssim.py).
"""
import os, glob, subprocess, sys
from pathlib import Path
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = Path('/c/Users/ryuji/oxi-main') if False else Path(os.path.abspath('.'))
RENDERER = ROOT / 'tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe'
SRC = ROOT / 'pipeline_data/golden_per_page'
WORD_PNG = ROOT / 'pipeline_data/per_page_ssim/word_png'
TMP = Path('c:/tmp/_s492ssim'); TMP.mkdir(parents=True, exist_ok=True)

STEMS = sys.argv[1:] or ['683ffcab86e2', '0e7af1ae8f21', 'd77a58485f16']


def render(docx, prefix, on):
    env = dict(os.environ); env.pop('OXI_S492_JCNATURAL', None)
    if on: env['OXI_S492_JCNATURAL'] = '1'
    r = subprocess.run([str(RENDERER), str(docx), prefix], capture_output=True, text=True, env=env)
    for line in (r.stdout + '\n' + r.stderr).splitlines():
        line = line.strip()
        if line.startswith('Saved ') and line.split(' (')[0].endswith('_p1.png'):
            return Path(line[len('Saved '):].split(' (')[0])
    # fall back to prefix_p1.png
    p = Path(prefix + '_p1.png')
    return p if p.exists() else None


def load_gray(p, size=None):
    im = Image.open(p).convert('L')
    if size and im.size != size:
        im = im.resize(size)
    return np.array(im)


def score(oxi_png, word_png):
    w = load_gray(word_png)
    o = load_gray(oxi_png, size=Image.open(word_png).size)
    return ssim(w, o)


tot_off = tot_on = 0.0
n = 0
for stem in STEMS:
    slices = sorted(glob.glob(str(SRC / (stem + '*_p*.docx'))))
    print('\n=== %s (%d slices) ===' % (stem, len(slices)))
    print('%-40s %7s %7s %8s' % ('slice', 'OFF', 'ON', 'delta'))
    for docx in slices:
        st = Path(docx).stem
        wpng = WORD_PNG / (st + '_p1.png')
        if not wpng.exists():
            # try without _p1
            cand = list(WORD_PNG.glob(st + '*.png'))
            if not cand:
                continue
            wpng = cand[0]
        off = render(docx, str(TMP / (st + '_off')), 0)
        on = render(docx, str(TMP / (st + '_on')), 1)
        if not off or not on:
            print('%-40s  render-fail' % st[:40]); continue
        try:
            so = score(off, wpng); sn = score(on, wpng)
        except Exception as e:
            print('%-40s  ssim-fail %s' % (st[:40], e)); continue
        d = sn - so
        flag = '' if abs(d) < 0.001 else ('  UP' if d > 0 else '  DOWN')
        print('%-40s %7.4f %7.4f %+8.4f%s' % (st[-40:], so, sn, d, flag))
        tot_off += so; tot_on += sn; n += 1

if n:
    print('\nNET over %d touched pages: OFF mean=%.4f  ON mean=%.4f  delta=%+.4f'
          % (n, tot_off / n, tot_on / n, (tot_on - tot_off) / n))
