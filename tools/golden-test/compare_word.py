#!/usr/bin/env python3
"""Compare Oxi renders against Word renders (SSIM)."""
import sys
from pathlib import Path

import cv2
import numpy as np
from skimage.metrics import structural_similarity as ssim

OUTPUT_DIR = Path(__file__).resolve().parent / "pixel_output"


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
    word_dir = OUTPUT_DIR / "word"
    oxi_dir = OUTPUT_DIR / "oxi"

    if not word_dir.exists():
        print("No Word renders found. Run render_word.py first.")
        sys.exit(1)

    scores = []
    for word_png in sorted(word_dir.glob("*.png")):
        oxi_png = oxi_dir / word_png.name
        if not oxi_png.exists():
            continue
        score = compare_images(oxi_png, word_png)
        if score is not None:
            scores.append((score, word_png.stem))

    scores.sort()
    for s, f in scores:
        print(f"  {s:.4f}  {f[:60]}")
    if scores:
        avg = sum(s for s, _ in scores) / len(scores)
        above_95 = sum(1 for s, _ in scores if s >= 0.95)
        above_90 = sum(1 for s, _ in scores if s >= 0.90)
        print(f"\nAverage SSIM: {avg:.4f} ({len(scores)} files)")
        print(f"  >= 0.95: {above_95}/{len(scores)}")
        print(f"  >= 0.90: {above_90}/{len(scores)}")


if __name__ == "__main__":
    main()
