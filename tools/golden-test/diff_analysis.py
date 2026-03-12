#!/usr/bin/env python3
"""Generate diff images and analyze WHERE Oxi vs LibreOffice diverge."""
import cv2
import numpy as np
from pathlib import Path

OUTPUT_DIR = Path(__file__).resolve().parent / "pixel_output"

def analyze_diff(name):
    oxi_path = OUTPUT_DIR / "oxi" / f"{name}.png"
    lo_path = OUTPUT_DIR / "libreoffice" / f"{name}.png"
    diff_dir = OUTPUT_DIR / "diff"
    diff_dir.mkdir(exist_ok=True)

    oxi = cv2.imread(str(oxi_path))
    lo = cv2.imread(str(lo_path))
    if oxi is None or lo is None:
        print(f"  SKIP {name}: missing image")
        return

    # Match sizes
    h = min(oxi.shape[0], lo.shape[0])
    w = min(oxi.shape[1], lo.shape[1])
    oxi = cv2.resize(oxi, (w, h))
    lo = cv2.resize(lo, (w, h))

    # Absolute difference
    diff = cv2.absdiff(oxi, lo)
    gray_diff = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)

    # Amplify for visibility
    amplified = cv2.normalize(gray_diff, None, 0, 255, cv2.NORM_MINMAX)

    # Create side-by-side comparison with diff
    # Top: Oxi, Middle: LibreOffice, Bottom: Diff (amplified)
    comparison = np.vstack([oxi, lo, cv2.cvtColor(amplified, cv2.COLOR_GRAY2BGR)])
    cv2.imwrite(str(diff_dir / f"{name}_comparison.png"), comparison)

    # Heatmap of differences
    heatmap = cv2.applyColorMap(amplified, cv2.COLORMAP_JET)
    # Blend heatmap with LO image for context
    blended = cv2.addWeighted(lo, 0.5, heatmap, 0.5, 0)
    cv2.imwrite(str(diff_dir / f"{name}_heatmap.png"), blended)

    # Analyze vertical drift: compute row-by-row difference
    row_diff = gray_diff.mean(axis=1)  # average diff per row

    # Find first row with significant difference (>10)
    significant = np.where(row_diff > 10)[0]
    first_diff_row = significant[0] if len(significant) > 0 else -1

    # Analyze horizontal vs vertical
    col_diff = gray_diff.mean(axis=0)
    total_diff = gray_diff.mean()

    # Split into quadrants
    mid_h, mid_w = h // 2, w // 2
    q_tl = gray_diff[:mid_h, :mid_w].mean()
    q_tr = gray_diff[:mid_h, mid_w:].mean()
    q_bl = gray_diff[mid_h:, :mid_w].mean()
    q_br = gray_diff[mid_h:, mid_w:].mean()

    # Split into horizontal bands (top 1/3, middle 1/3, bottom 1/3)
    h3 = h // 3
    band_top = gray_diff[:h3, :].mean()
    band_mid = gray_diff[h3:2*h3, :].mean()
    band_bot = gray_diff[2*h3:, :].mean()

    print(f"\n{'='*60}")
    print(f"  {name}")
    print(f"{'='*60}")
    print(f"  Image size: {w}x{h}")
    print(f"  Overall mean diff: {total_diff:.2f}")
    print(f"  First significant diff at row: {first_diff_row} ({first_diff_row*100/h:.1f}% from top)")
    print(f"  Quadrants (TL/TR/BL/BR): {q_tl:.1f} / {q_tr:.1f} / {q_bl:.1f} / {q_br:.1f}")
    print(f"  Horizontal bands (top/mid/bot): {band_top:.1f} / {band_mid:.1f} / {band_bot:.1f}")

    # Detect if the difference is mainly a vertical shift
    # Try shifting oxi image up/down and see if SSIM improves
    best_shift = 0
    best_score = total_diff
    for shift in range(-30, 31):
        if shift == 0:
            continue
        if shift > 0:
            shifted_oxi = oxi[shift:, :]
            shifted_lo = lo[:h-shift, :]
        else:
            shifted_oxi = oxi[:h+shift, :]
            shifted_lo = lo[-shift:, :]
        d = cv2.absdiff(shifted_oxi, shifted_lo)
        score = cv2.cvtColor(d, cv2.COLOR_BGR2GRAY).mean()
        if score < best_score:
            best_score = score
            best_shift = shift

    if best_shift != 0:
        print(f"  ** Optimal vertical shift: {best_shift}px (reduces diff from {total_diff:.2f} to {best_score:.2f})")
        # At 150 DPI, 1px ≈ 0.48pt
        print(f"     = ~{best_shift * 0.48:.1f}pt vertical offset")
    else:
        print(f"  No vertical shift helps")


def main():
    # Analyze worst-scoring files
    targets = [
        "683ffcab86e2_20230331_resources_open_data_contract_addon_00",
        "1ec1091177b1_006",
        "a1d6e4efa2e7_tokumei_08_01-4",
        "2ea81a8441cc_0025006-192",
        "6514f214e482_tokumei_08_01-2",
        "1636d28e2c46_tokumei_08_04",
        "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00",
        "a47e6c6b2ca1_order_08",
        "15076df085f5_tokumei_08_09",
        "459f05f1e877_kyodokenkyuyoushiki01",
    ]

    for name in targets:
        analyze_diff(name)

if __name__ == "__main__":
    main()
