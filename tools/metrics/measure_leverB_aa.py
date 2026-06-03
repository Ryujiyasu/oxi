# -*- coding: utf-8 -*-
"""S492t — Lever B (renderer AA cap). On dense-CJK pages, render Oxi DWrite at
supersample 1/2/4, SSIM vs cached Word PNG; also compare GRAY-LEVEL HISTOGRAMS
(Oxi vs Word) to diagnose: is the cap stroke-EDGE AA (supersample helps) or CORE
darkness (supersample won't)? S460 chose 1x for perf; no-ROI removes that. cp932-safe.
"""
import os, glob, subprocess, json
from pathlib import Path
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

ROOT = Path(os.path.abspath('.'))
RENDER = ROOT / 'tools/oxi-dwrite-renderer/target/release/oxi-dwrite-renderer.exe'
targets = json.load(open('c:/tmp/bottomN/targets.json', encoding='utf-8'))
# dense-CJK representative pages (the AA-capped ones per scout)
SEL = {('683ffcab86e2', 2), ('b35123fe8efc', 1), ('d77a58485f16', 9),
       ('15076df085f5', 1), ('29dc6e8943fe', 5)}
sel = [t for t in targets if (t['doc_id'], t['page']) in SEL]

def gray(p, size=None):
    im = Image.open(p).convert('L')
    if size and im.size != size: im = im.resize(size)
    return np.array(im)

def render(docx, doc_id, ss):
    pre = 'c:/tmp/_b_%s_ss%d' % (doc_id, ss)
    subprocess.run([str(RENDER), docx, pre, '150', '--supersample=%d' % ss],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return pre

def darkfrac(a):
    # fraction of near-black core pixels (<64) and mid-gray (64..200)
    tot = a.size
    return (a < 64).sum() / tot, ((a >= 64) & (a < 200)).sum() / tot

print("doc            p   ss1     ss2     ss4   | core<64 (Word/ss1/ss4)  mid-gray (Word/ss1/ss4)")
for t in sel:
    if not (t['word_png'] and os.path.exists(t['word_png'])): continue
    sz = Image.open(t['word_png']).size; w = gray(t['word_png'])
    wc, wm = darkfrac(w)
    row = []
    dk = []
    for ss in (1, 2, 4):
        pre = render(t['docx'], t['doc_id'], ss)
        p = pre + ('_p%d.png' % t['page'])
        if not os.path.exists(p):
            row.append(None); dk.append((None, None)); continue
        g = gray(p, sz)
        row.append(ssim(w, g))
        dk.append(darkfrac(g))
    def f(x): return '%.4f' % x if x is not None else '  -  '
    c1 = dk[0]; c4 = dk[2]
    print("%-14s p%-2d %s %s %s | %.3f/%.3f/%.3f      %.3f/%.3f/%.3f" % (
        t['doc_id'], t['page'], f(row[0]), f(row[1]), f(row[2]),
        wc, c1[0] if c1[0] is not None else 0, c4[0] if c4[0] is not None else 0,
        wm, c1[1] if c1[1] is not None else 0, c4[1] if c4[1] is not None else 0))
