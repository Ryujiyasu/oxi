#!/usr/bin/env python3
"""Quick SSIM comparison: Oxi PDF vs LibreOffice PDF, per page.
Caches LO PNGs to speed up repeated comparisons."""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import cv2
import fitz  # PyMuPDF
import numpy as np
from skimage.metrics import structural_similarity as ssim

SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
DPI = 150
DOCX = Path(__file__).resolve().parent.parent.parent / "tests" / "fixtures" / "comprehensive_test.docx"
OXI_PDF = Path(__file__).resolve().parent / "pixel_output" / "oxi_output.pdf"
OUTPUT_DIR = Path(__file__).resolve().parent / "pixel_output"


def pdf_to_pngs(pdf_path, prefix, dpi=DPI):
    doc = fitz.open(str(pdf_path))
    paths = []
    for i, page in enumerate(doc):
        zoom = dpi / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        out = OUTPUT_DIR / f"{prefix}_page{i+1}.png"
        pix.save(str(out))
        paths.append(out)
    doc.close()
    return paths


def get_lo_pngs():
    """Return LO reference PNGs, using cache if available."""
    # Check for cached LO PNGs
    cached = sorted(OUTPUT_DIR.glob("lo_page*.png"))
    if cached:
        return cached
    # Render from scratch
    tmp = tempfile.mkdtemp(prefix="lo_")
    profile = tmp + "/profile"
    pdf_dir = tmp + "/pdf"
    import os; os.makedirs(profile); os.makedirs(pdf_dir)
    subprocess.run(
        [SOFFICE, "--headless", "--norestore", "--nologo",
         "--convert-to", "pdf", "--outdir", pdf_dir,
         f"-env:UserInstallation=file:///{profile.replace(chr(92), '/')}",
         str(DOCX)],
        capture_output=True, timeout=60,
    )
    pdf = Path(pdf_dir) / (DOCX.stem + ".pdf")
    pngs = pdf_to_pngs(pdf, "lo") if pdf.exists() else []
    shutil.rmtree(tmp, ignore_errors=True)
    return pngs


def compare(img1_path, img2_path):
    img1 = cv2.imread(str(img1_path))
    img2 = cv2.imread(str(img2_path))
    h = min(img1.shape[0], img2.shape[0])
    w = min(img1.shape[1], img2.shape[1])
    img1 = cv2.resize(img1, (w, h))
    img2 = cv2.resize(img2, (w, h))
    g1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
    g2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
    score, diff_map = ssim(g1, g2, full=True)
    diff_vis = (255 * (1 - diff_map)).clip(0, 255).astype(np.uint8)
    return score, diff_vis


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Render Oxi PDF to PNGs
    oxi_pngs = pdf_to_pngs(OXI_PDF, "oxi")

    # Get LO reference PNGs (cached)
    lo_pngs = get_lo_pngs()

    # Compare
    n = min(len(oxi_pngs), len(lo_pngs))
    scores = []
    for i in range(n):
        score, diff_vis = compare(oxi_pngs[i], lo_pngs[i])
        scores.append(score)
        diff_path = OUTPUT_DIR / f"diff_page{i+1}.png"
        cv2.imwrite(str(diff_path), diff_vis)
        print(f"P{i+1}: {score:.4f}", end="  ")

    if scores:
        avg = sum(scores) / len(scores)
        print(f"Avg: {avg:.4f}")


if __name__ == "__main__":
    main()
