#!/usr/bin/env python3
"""Quick re-render Oxi and compare against existing LibreOffice images."""
import json
import os
import sys
import time
from pathlib import Path

import cv2
import numpy as np
from skimage.metrics import structural_similarity as ssim

OUTPUT_DIR = Path(__file__).resolve().parent / "pixel_output"
WEB_ROOT = Path(__file__).resolve().parent.parent.parent
PRINT_PREVIEW_URL = "http://localhost:8080/web/print-preview.html"


def compare_images(img1_path, img2_path):
    img1 = cv2.imread(str(img1_path))
    img2 = cv2.imread(str(img2_path))
    if img1 is None or img2 is None:
        return None
    h = min(img1.shape[0], img2.shape[0])
    w = min(img1.shape[1], img2.shape[1])
    img1 = cv2.resize(img1, (w, h))
    img2 = cv2.resize(img2, (w, h))
    gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
    gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
    score, _ = ssim(gray1, gray2, full=True)
    return round(score, 4)


def main():
    from playwright.sync_api import sync_playwright

    # Find all docx files that have LibreOffice renders
    lo_dir = OUTPUT_DIR / "libreoffice"
    oxi_dir = OUTPUT_DIR / "oxi"
    oxi_dir.mkdir(parents=True, exist_ok=True)

    # Map LO PNGs back to docx files
    fixtures_dir = WEB_ROOT / "tests" / "fixtures"
    golden_dir = Path(__file__).resolve().parent / "documents" / "docx"

    test_files = []
    for f in sorted(fixtures_dir.glob("*.docx")):
        if (lo_dir / f"{f.stem}.png").exists():
            test_files.append(f)
    if golden_dir.exists():
        for f in sorted(golden_dir.glob("*.docx"))[:30]:
            if (lo_dir / f"{f.stem}.png").exists():
                test_files.append(f)

    total = len(test_files)
    print(f"Re-rendering {total} files with Oxi...", flush=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1280, "height": 2000})

        for i, docx_path in enumerate(test_files):
            out_png = oxi_dir / f"{docx_path.stem}.png"
            try:
                page = context.new_page()
                page.goto(PRINT_PREVIEW_URL, wait_until="networkidle", timeout=30000)
                page.wait_for_function("() => window.__oxiWasmReady === true", timeout=15000)
                page.locator("#fileInput").set_input_files(str(docx_path))
                page.wait_for_function("() => window.__oxiRenderDone === true", timeout=30000)
                page.wait_for_timeout(500)
                canvas = page.locator("#page-0")
                if canvas.count() > 0:
                    canvas.screenshot(path=str(out_png))
                page.close()
            except Exception as e:
                print(f"  ERROR {docx_path.name}: {e}", flush=True)
                try:
                    page.close()
                except:
                    pass

        browser.close()

    # Compare
    scores = []
    for docx_path in test_files:
        oxi_png = oxi_dir / f"{docx_path.stem}.png"
        lo_png = lo_dir / f"{docx_path.stem}.png"
        if oxi_png.exists() and lo_png.exists():
            score = compare_images(oxi_png, lo_png)
            if score is not None:
                scores.append((score, docx_path.name))

    scores.sort()
    for s, f in scores:
        print(f"  {s:.4f}  {f}")
    if scores:
        avg = sum(s for s, _ in scores) / len(scores)
        print(f"\nAverage SSIM: {avg:.4f} ({len(scores)} files)")


if __name__ == "__main__":
    main()
