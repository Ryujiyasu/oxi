# -*- coding: utf-8 -*-
"""Word | Oxi | difference, for one page of one document -- the three panels a
layout bug is read from.  The difference panel paints Word-only ink red and
Oxi-only ink blue, so a shifted line reads as a red/blue pair and a missing one
as a solid red block.

Usage: python tools/metrics/_worddiff_panel.py <base> <page> [out.png] [y0 y1]
  <base> is the pipeline_data/word_png directory name (prefix is enough);
  y0/y1 crop the page in POINTS so a band can be inspected close up.
"""
import glob
import os
import re
import subprocess
import sys
import tempfile

import numpy as np
from PIL import Image, ImageDraw, ImageFont

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, REPO)
from pipeline.config import RENDER_DPI  # noqa: E402
from pipeline.ssim_calculator import _load_rgb, _resize_to_match  # noqa: E402
from skimage.metrics import structural_similarity as ssim  # noqa: E402

base, page = sys.argv[1], int(sys.argv[2])
out = sys.argv[3] if len(sys.argv) > 3 else os.path.join(REPO, "pipeline_data", "worddiff_%s_p%d.png" % (base[:12], page))
crop = (float(sys.argv[4]), float(sys.argv[5])) if len(sys.argv) > 5 else None
wdir = glob.glob(os.path.join(REPO, "pipeline_data", "word_png", base + "*"))[0]
word_png = os.path.join(wdir, "page_%04d.png" % page)
docx = glob.glob(os.path.join(REPO, "tools", "golden-test", "documents", "docx", os.path.basename(wdir).split("_")[0] + "*.docx"))[0]
DW = os.environ.get("OXI_DWRITE_EXE") or os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
tmp = tempfile.mkdtemp(prefix="wdp_")
env = dict(os.environ)
subprocess.run([DW, os.path.abspath(docx), os.path.join(tmp, "p"), str(RENDER_DPI)], capture_output=True, timeout=900, env=env)
pages = sorted(glob.glob(os.path.join(tmp, "p*.png")), key=lambda f: int(re.search(r"(\d+)\.png$", f).group(1)))
oxi_png = pages[page - 1]

w = _load_rgb(word_png)
o = _resize_to_match(_load_rgb(oxi_png), w)
print("page %d  SSIM %.4f  (%s)" % (page, float(ssim(w, o, channel_axis=2, data_range=255)), os.path.basename(wdir)))

wi = Image.fromarray(w.astype("uint8"))
oi = Image.fromarray(o.astype("uint8"))
if crop:
    s = wi.height / 842.0
    box = (0, int(crop[0] * s), wi.width, int(crop[1] * s))
    wi, oi = wi.crop(box), oi.crop(box)
wg = np.array(wi.convert("L")).astype(int)
og = np.array(oi.convert("L")).astype(int)
diff = np.full(wg.shape + (3,), 255, dtype="uint8")
word_only = (wg < 160) & (og >= 160)
oxi_only = (og < 160) & (wg >= 160)
both = (wg < 160) & (og < 160)
diff[both] = (185, 190, 200)
diff[word_only] = (200, 40, 40)
diff[oxi_only] = (40, 80, 210)
di = Image.fromarray(diff)
print("   ink: Word-only %d px, Oxi-only %d px, shared %d px" % (word_only.sum(), oxi_only.sum(), both.sum()))

H = int(os.environ.get("PANEL_H", "1500"))
panels = [("Word", wi), ("Oxi", oi), ("difference (red = Word only, blue = Oxi only)", di)]
imgs = [(lab, im.resize((int(im.width * H / im.height), H), Image.LANCZOS)) for lab, im in panels]
pad, top = 20, 44
W = sum(im.width for _, im in imgs) + pad * (len(imgs) + 1)
canvas = Image.new("RGB", (W, H + top + pad), (247, 247, 244))
d = ImageDraw.Draw(canvas)
try:
    font = ImageFont.truetype("C:/Windows/Fonts/segoeui.ttf", 26)
except Exception:
    font = ImageFont.load_default()
x = pad
for lab, im in imgs:
    d.text((x, 12), lab, fill=(28, 34, 48), font=font)
    canvas.paste(im, (x, top))
    d.rectangle([x - 1, top - 1, x + im.width, top + H], outline=(214, 217, 224))
    x += im.width + pad
canvas.save(out)
print("wrote", out, canvas.size)
