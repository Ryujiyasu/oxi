# -*- coding: utf-8 -*-
"""S492r — evaluate the existing S467 VSNAP (OXI_S467_VSNAP) = Lever C (exact line-height
+ 0.75pt top-snap) on the C-affected bottom-N pages. Render Oxi DWrite OFF vs S467,
SSIM vs cached Word PNG. cp932-safe (UTF-8 file, ASCII out)."""
import os, json, subprocess
from pathlib import Path
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = Path(os.path.abspath('.'))
RENDER = ROOT / 'tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe'
targets = json.load(open('c:/tmp/bottomN/targets.json', encoding='utf-8'))
# C-affected pages (per scout) + 683f p2 as no-regress sentinel
C_PAGES = {('a1d6e4efa2e7', 4), ('e3c545fac7a7', 4), ('e3c545fac7a7', 11),
           ('15076df085f5', 1), ('683ffcab86e2', 2), ('d4d126dfe1d9', 4)}
sel = [t for t in targets if (t['doc_id'], t['page']) in C_PAGES]

def gray(p, size=None):
    im = Image.open(p).convert('L')
    if size and im.size != size: im = im.resize(size)
    return np.array(im)

def render(docx, doc_id, vsnap):
    env = dict(os.environ); env.pop('OXI_S467_VSNAP', None)
    if vsnap: env['OXI_S467_VSNAP'] = '1'
    prefix = 'c:/tmp/_s467_%s_%d' % (doc_id, vsnap)
    subprocess.run([str(RENDER), docx, prefix, '150'], stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL, env=env)
    return prefix

print("doc            p   ssim_OFF  ssim_S467   delta")
tot = 0.0; n = 0
for t in sel:
    if not (t['word_png'] and os.path.exists(t['word_png'])):
        continue
    w = gray(t['word_png']); sz = Image.open(t['word_png']).size
    off_prefix = render(t['docx'], t['doc_id'], 0)
    on_prefix = render(t['docx'], t['doc_id'], 1)
    off_png = off_prefix + ('_p%d.png' % t['page'])
    on_png = on_prefix + ('_p%d.png' % t['page'])
    if not (os.path.exists(off_png) and os.path.exists(on_png)):
        print("%-14s p%-2d  render-miss" % (t['doc_id'], t['page'])); continue
    so = ssim(w, gray(off_png, sz)); sn = ssim(w, gray(on_png, sz))
    d = sn - so
    flag = '  UP' if d > 0.002 else ('  DOWN' if d < -0.002 else '')
    print("%-14s p%-2d  %.4f    %.4f    %+.4f%s" % (t['doc_id'], t['page'], so, sn, d, flag))
    tot += d; n += 1
if n:
    print("\nmean SSIM delta (S467 - OFF) over %d C-pages: %+.4f" % (n, tot / n))
