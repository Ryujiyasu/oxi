# -*- coding: utf-8 -*-
"""Crop a row range from word_png and the fresh oxi render, stack vertically for visual compare."""
import os, sys, glob
from PIL import Image
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
WORD_PNG = os.path.join(ROOT, 'pipeline_data', 'word_png')

stem = sys.argv[1]
y0 = int(sys.argv[2]); y1 = int(sys.argv[3])
page = int(sys.argv[4]) if len(sys.argv) > 4 else 1
oxi = os.path.join('c:/tmp', 's531_' + stem[:12] + ('_p%d.png' % page))
wdir = glob.glob(os.path.join(WORD_PNG, stem + '*'))[0]
wpng = os.path.join(wdir, 'page_%04d.png' % page)
wim = Image.open(wpng).convert('RGB')
oim = Image.open(oxi).convert('RGB')
W = max(wim.width, oim.width)
wc = wim.crop((0, y0, wim.width, y1))
oc = oim.crop((0, y0, oim.width, y1))
h = (y1 - y0)
canvas = Image.new('RGB', (W, h * 2 + 6), (255, 0, 0))
canvas.paste(wc, (0, 0))
canvas.paste(oc, (0, h + 6))
out = 'c:/tmp/_s531_crop.png'
canvas.save(out)
print('WORD (top) vs OXI (bottom), rows %d-%d -> %s  size=%s' % (y0, y1, out, canvas.size))
