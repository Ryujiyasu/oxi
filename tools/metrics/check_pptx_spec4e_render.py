# -*- coding: utf-8 -*-
"""Spec #4e render-side check: render Word PDF pages with fitz at 150 DPI
and compare against Oxi PNGs via pixel SSIM (per-slide)."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import fitz
from skimage.metrics import structural_similarity as ssim
import numpy as np
from PIL import Image

base = r"pipeline_data\pptx_probes\spec4e_multifont"
pdf_path = os.path.join(base, "spec4e_multifont.pdf")

doc = fitz.open(pdf_path)
print(f"Word PDF pages: {doc.page_count}")

for i in range(doc.page_count):
    page = doc.load_page(i)
    # Render at 150 DPI
    pix = page.get_pixmap(dpi=150)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = img[:, :, :3]
    img = img.astype(np.float64)

    oxi_path = os.path.join(base, f"oxi_s{i+1}.png")
    if not os.path.exists(oxi_path):
        print(f"slide {i+1}: Oxi PNG missing")
        continue
    oxi_img = np.asarray(Image.open(oxi_path).convert("RGB")).astype(np.float64)

    # Align sizes (Oxi PNG is supersample=2 -> 2x of the 150dpi render)
    h = min(img.shape[0], oxi_img.shape[0])
    w = min(img.shape[1], oxi_img.shape[1])
    a = img[:h, :w]
    b = oxi_img[:h, :w]
    # Downsample Oxi to the 150dpi size for a fair comparison
    from skimage.transform import resize
    b_small = resize(b, (pix.height, pix.width), preserve_range=True, anti_aliasing=True)
    sc = ssim(a / 255.0, b_small / 255.0, channel_axis=2, data_range=1.0)
    print(f"slide {i+1}: WordPDF-vs-Oxi SSIM = {sc:.4f}  ({pix.width}x{pix.height})")
doc.close()
