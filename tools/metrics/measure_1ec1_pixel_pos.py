"""Pixel-measure □ bullet x positions in Word vs Oxi PNGs.

Word PNG: pipeline_data/word_png/1ec1091177b1_006/page_0001.png
Oxi PNG:  pipeline_data/oxi_png/1ec1091177b1_006/page_0001.png

PNG width is page_w_pt * dpi/72. Default 96 DPI → 1pt ≈ 1.333 px.
Find the x of left-edge of dark pixels at known y-positions of bullets.

Bullet y-positions (from layout JSON):
  □1: y=178pt — TextBox 1 (top body region)
  □3: y=350pt — TextBox 4 (bottom body region with 5.25pt indent)

For each y, find leftmost dark pixel near body margin (~30-60pt range).
"""
import os
import sys
from PIL import Image

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WORD_PNG = "pipeline_data/word_png/1ec1091177b1_006/page_0001.png"
OXI_PNG = "pipeline_data/oxi_png/1ec1091177b1_006/page_0001.png"


def measure_left_edge(png_path, y_pt, x_min_pt, x_max_pt, label):
    """For the given y range, scan x_min..x_max for first dark pixel."""
    img = Image.open(png_path).convert("L")
    w, h = img.size
    print(f"\n{label}: {png_path} — {w}x{h}")
    # PNG width in px = page_w_pt × dpi/72. Page 595.30pt.
    # Calibrate dpi: dpi = w × 72 / 595.30
    page_w_pt = 595.30
    dpi = w * 72.0 / page_w_pt
    print(f"  Calibrated DPI: {dpi:.2f}, 1pt = {w/page_w_pt:.4f} px")
    pix = img.load()
    # y_pt range ±5pt (line height)
    for y_target in [y_pt]:
        ymin_px = int((y_target - 8) * h / 841.9)
        ymax_px = int((y_target + 8) * h / 841.9)
        xmin_px = int(x_min_pt * w / page_w_pt)
        xmax_px = int(x_max_pt * w / page_w_pt)
        print(f"  Scanning y_pt={y_target} (px {ymin_px}-{ymax_px}), "
              f"x_pt={x_min_pt}-{x_max_pt} (px {xmin_px}-{xmax_px})")
        first_dark_x = None
        for x_px in range(xmin_px, xmax_px):
            for y_px in range(ymin_px, ymax_px):
                if pix[x_px, y_px] < 100:  # dark
                    first_dark_x = x_px
                    break
            if first_dark_x is not None:
                break
        if first_dark_x is None:
            print(f"    No dark pixel found in range")
        else:
            x_pt = first_dark_x * page_w_pt / w
            print(f"    First dark x = {first_dark_x}px = {x_pt:.2f}pt")
            return x_pt
    return None


def main():
    # □1 at y≈178pt, □3 at y≈350pt
    # Body text margin ~ 42.55pt; scan 30-65pt range
    print("=" * 60)
    print("□1 (TB[1] — should be similar between Word and Oxi)")
    word_b1 = measure_left_edge(WORD_PNG, 178, 30, 65, "Word")
    oxi_b1 = measure_left_edge(OXI_PNG, 178, 30, 65, "Oxi")
    if word_b1 and oxi_b1:
        print(f"\n  □1 diff: Oxi - Word = {oxi_b1 - word_b1:+.2f}pt")

    print("\n" + "=" * 60)
    print("□3 (TB[4] — user reports +1.81pt offset in Oxi)")
    word_b3 = measure_left_edge(WORD_PNG, 350, 30, 65, "Word")
    oxi_b3 = measure_left_edge(OXI_PNG, 350, 30, 65, "Oxi")
    if word_b3 and oxi_b3:
        print(f"\n  □3 diff: Oxi - Word = {oxi_b3 - word_b3:+.2f}pt")

    print("\n" + "=" * 60)
    print("Body margin reference: scan body text 'の' at y≈92pt")
    word_body = measure_left_edge(WORD_PNG, 92, 30, 65, "Word")
    oxi_body = measure_left_edge(OXI_PNG, 92, 30, 65, "Oxi")
    if word_body and oxi_body:
        print(f"\n  Body diff: Oxi - Word = {oxi_body - word_body:+.2f}pt")


if __name__ == "__main__":
    main()
