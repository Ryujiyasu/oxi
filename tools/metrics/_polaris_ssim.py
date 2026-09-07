# -*- coding: utf-8 -*-
"""Score a Polaris Office PDF export against the Word reference PNGs, with the
same RGB SSIM the pipeline gate uses, and print LibreOffice / Oxi beside it when
those renders exist.

Usage: python tools/metrics/_polaris_ssim.py <polaris.pdf> <word_png dir name>
"""
import glob
import os
import re
import subprocess
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, REPO)
from pipeline.config import RENDER_DPI  # noqa: E402
from pipeline.ssim_calculator import _load_rgb, _resize_to_match  # noqa: E402
from skimage.metrics import structural_similarity as ssim  # noqa: E402
import fitz  # noqa: E402

pdf, name = sys.argv[1], sys.argv[2]
wdir = glob.glob(os.path.join(REPO, "pipeline_data", "word_png", name + "*"))[0]
wpages = sorted(glob.glob(os.path.join(wdir, "page_*.png")))
tmp = tempfile.mkdtemp(prefix="polaris_")
doc = fitz.open(pdf)
ppages = []
for i in range(len(doc)):
    p = os.path.join(tmp, "p%03d.png" % (i + 1))
    doc[i].get_pixmap(dpi=RENDER_DPI).save(p)
    ppages.append(p)


def score(a_path, b_path):
    a = _load_rgb(a_path)
    b = _resize_to_match(_load_rgb(b_path), a)
    return float(ssim(a, b, channel_axis=2, data_range=255))


others = {}
lib = os.path.join(REPO, "pipeline_data", "libra_png", os.path.basename(wdir))
if os.path.isdir(lib):
    others["LibreOffice"] = sorted(glob.glob(os.path.join(lib, "page_*.png")))
DW = os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release", "oxi-dwrite-renderer.exe")
docx = glob.glob(os.path.join(REPO, "tools", "golden-test", "documents", "docx", name.split("_")[0] + "*.docx"))
if docx and os.path.exists(DW):
    otmp = tempfile.mkdtemp(prefix="oxi_")
    subprocess.run([DW, docx[0], os.path.join(otmp, "p"), str(RENDER_DPI)], capture_output=True, timeout=900)
    others["Oxi"] = sorted(glob.glob(os.path.join(otmp, "p*.png")), key=lambda f: int(re.search(r"(\d+)\.png$", f).group(1)))

print("%s | Word %d pages, Polaris %d pages" % (os.path.basename(wdir), len(wpages), len(ppages)))
rows = {"Polaris": ppages}
rows.update(others)
for label, pages in rows.items():
    sc = [score(w, pages[i]) for i, w in enumerate(wpages) if i < len(pages)]
    if sc:
        print("  %-12s pages=%-3d mean=%.4f  %s" % (label, len(pages), sum(sc) / len(sc), " ".join("%.3f" % s for s in sc[:10])))
    else:
        print("  %-12s no pages" % label)
