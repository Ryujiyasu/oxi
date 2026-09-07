# -*- coding: utf-8 -*-
"""Three-panel comparison of one page: Word (ground truth PNG) | LibreOffice | Oxi
(DWrite render), each labelled with its SSIM against Word, for slides and notes.

Usage: python tools/metrics/_compare3.py <doc-dir-name> <page> [out.png]
  <doc-dir-name> is the pipeline_data/word_png/<name> directory (the docx is looked up
  under tools/golden-test/documents/docx by the same stem).
"""
import glob
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
from pipeline.config import RENDER_DPI  # noqa: E402
from pipeline.ssim_calculator import _load_rgb, _resize_to_match  # noqa: E402
from skimage.metrics import structural_similarity as ssim  # noqa: E402

DW = os.environ.get("OXI_DWRITE_EXE") or str(REPO / "tools" / "oxi-dwrite-renderer" / "target" / "release" / "oxi-dwrite-renderer.exe")

name, page = sys.argv[1], int(sys.argv[2])
out = sys.argv[3] if len(sys.argv) > 3 else str(REPO / "pipeline_data" / "compare3_%s_p%d.png" % (name[:12], page))
word_png = REPO / "pipeline_data" / "word_png" / name / ("page_%04d.png" % page)
libra_png = REPO / "pipeline_data" / "libra_png" / name / ("page_%04d.png" % page)
docx = glob.glob(str(REPO / "tools" / "golden-test" / "documents" / "docx" / (name + ".docx")))
if not docx:
    docx = glob.glob(str(REPO / "tools" / "golden-test" / "documents" / "docx" / (name.split("_")[0] + "*.docx")))
docx = docx[0]
tmp = tempfile.mkdtemp(prefix="cmp3_")
subprocess.run([DW, docx, os.path.join(tmp, "p"), str(RENDER_DPI)], capture_output=True, timeout=600)
import re
oxi_pages = sorted(glob.glob(os.path.join(tmp, "p*.png")), key=lambda f: int(re.search(r"(\d+)\.png$", f).group(1)))
oxi_png = oxi_pages[page - 1] if page - 1 < len(oxi_pages) else None
print("oxi pages rendered:", len(oxi_pages), "| using", oxi_png)


def score(target):
    a = _load_rgb(str(word_png))
    b = _load_rgb(str(target))
    b = _resize_to_match(b, a)
    return float(ssim(a, b, channel_axis=2, data_range=255))


panels = [("Word (Microsoft 365)", str(word_png), None)]
if libra_png.exists():
    panels.append(("LibreOffice", str(libra_png), score(libra_png)))
if oxi_png:
    panels.append(("Oxi", oxi_png, score(oxi_png)))

H = 1400
imgs = []
for label, path, sc in panels:
    im = Image.open(path).convert("RGB")
    w = int(im.width * H / im.height)
    imgs.append((label, im.resize((w, H), Image.LANCZOS), sc))
pad, top = 24, 96
W = sum(im.width for _, im, _ in imgs) + pad * (len(imgs) + 1)
canvas = Image.new("RGB", (W, H + top + pad), (247, 247, 244))
draw = ImageDraw.Draw(canvas)
try:
    font = ImageFont.truetype("C:/Windows/Fonts/segoeui.ttf", 34)
    font_s = ImageFont.truetype("C:/Windows/Fonts/segoeui.ttf", 26)
except Exception:
    font = font_s = ImageFont.load_default()
x = pad
for label, im, sc in imgs:
    draw.text((x, 18), label, fill=(28, 34, 48), font=font)
    if sc is not None:
        draw.text((x, 58), "similarity to Word: %.3f" % sc, fill=(91, 99, 117) if sc < 0.95 else (43, 63, 140), font=font_s)
    else:
        draw.text((x, 58), "ground truth", fill=(91, 99, 117), font=font_s)
    canvas.paste(im, (x, top))
    draw.rectangle([x - 1, top - 1, x + im.width, top + H], outline=(217, 220, 227))
    x += im.width + pad
canvas.save(out)
print("wrote", out, canvas.size, [(l, round(s, 3) if s else None) for l, _, s in imgs])
