#!/usr/bin/env python3
"""Render all golden test docx files with Oxi (Playwright) and compare against Word renders."""
import sys
import time
from pathlib import Path

import cv2
import numpy as np
from skimage.metrics import structural_similarity as ssim

WEB_ROOT = Path(__file__).resolve().parent.parent.parent
PRINT_PREVIEW_URL = "http://localhost:8080/web/print-preview.html"
OUTPUT_DIR = Path(__file__).resolve().parent / "pixel_output"
DPI = 150


def render_oxi(docx_files, output_dir):
    from playwright.sync_api import sync_playwright
    import subprocess, urllib.request

    oxi_dir = output_dir / "oxi"
    oxi_dir.mkdir(parents=True, exist_ok=True)

    server_process = None
    try:
        try:
            urllib.request.urlopen("http://localhost:8080/web/", timeout=3)
            print("  Using existing server on :8080")
        except Exception:
            server_process = subprocess.Popen(
                [sys.executable, "-m", "http.server", "8080"],
                cwd=str(WEB_ROOT),
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            time.sleep(2)
            print("  Started server on :8080")

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={"width": 1280, "height": 2000})
            ok = 0
            total = len(docx_files)

            for i, docx_path in enumerate(docx_files):
                stem = docx_path.stem
                out_png = oxi_dir / f"{stem}.png"
                print(f"  [{i+1}/{total}] {stem[:50]}...", end=" ", flush=True)

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
                        print("OK")
                        ok += 1
                    else:
                        print("NO CANVAS")
                except Exception as e:
                    print(f"ERROR: {str(e)[:80]}")
                finally:
                    page.close()

            browser.close()
            print(f"\n  Rendered: {ok}/{total}")
    finally:
        if server_process:
            server_process.terminate()


def compare_word(output_dir):
    word_dir = output_dir / "word"
    oxi_dir = output_dir / "oxi"

    if not word_dir.exists():
        print("No Word renders found!")
        return

    scores = []
    for word_png in sorted(word_dir.glob("*.png")):
        oxi_png = oxi_dir / word_png.name
        if not oxi_png.exists():
            continue
        img1 = cv2.imread(str(oxi_png))
        img2 = cv2.imread(str(word_png))
        if img1 is None or img2 is None:
            continue
        h = min(img1.shape[0], img2.shape[0])
        w = min(img1.shape[1], img2.shape[1])
        img1 = cv2.resize(img1, (w, h))
        img2 = cv2.resize(img2, (w, h))
        g1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
        g2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
        score, _ = ssim(g1, g2, full=True)
        scores.append((round(score, 4), word_png.stem))

    scores.sort()
    for s, f in scores:
        print(f"  {s:.4f}  {f[:60]}")
    if scores:
        avg = sum(s for s, _ in scores) / len(scores)
        above_98 = sum(1 for s, _ in scores if s >= 0.98)
        above_95 = sum(1 for s, _ in scores if s >= 0.95)
        above_90 = sum(1 for s, _ in scores if s >= 0.90)
        print(f"\nAverage SSIM: {avg:.4f} ({len(scores)} files)")
        print(f"  >= 0.98: {above_98}/{len(scores)}")
        print(f"  >= 0.95: {above_95}/{len(scores)}")
        print(f"  >= 0.90: {above_90}/{len(scores)}")


def main():
    fixtures_dir = WEB_ROOT / "tests" / "fixtures"
    golden_dir = Path(__file__).resolve().parent / "documents" / "docx"

    # Collect files that have Word renders
    word_dir = OUTPUT_DIR / "word"
    word_stems = {p.stem for p in word_dir.glob("*.png")} if word_dir.exists() else set()

    test_files = []
    for f in sorted(fixtures_dir.glob("*.docx")):
        if f.stem in word_stems:
            test_files.append(f)
    if golden_dir.exists():
        for f in sorted(golden_dir.glob("*.docx")):
            if f.stem in word_stems:
                test_files.append(f)

    print(f"=== Rendering {len(test_files)} files with Oxi ===")
    render_oxi(test_files, OUTPUT_DIR)

    print(f"\n=== Comparing Oxi vs Word (SSIM) ===")
    compare_word(OUTPUT_DIR)


if __name__ == "__main__":
    main()
