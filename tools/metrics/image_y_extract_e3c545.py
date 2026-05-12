"""Extract text Y positions from Word PNG via image processing.

Approach: find rows with text content (rows where dark pixels exist
within the body width), cluster contiguous text rows into "line bands",
report each band's top Y in pt.

Then compare to Oxi PNG to find true Y-position divergence.

Uses Pillow.
"""
from __future__ import annotations
import sys
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).parent.parent.parent
WORD_P2 = ROOT / "pipeline_data/word_png/e3c545fac7a7_LOD_Handbook/page_0002.png"
OXI_P2 = ROOT / "pipeline_data/oxi_png/e3c545fac7a7_LOD_Handbook/page_0002.png"


def find_text_rows(img: Image.Image, x_start_px: int, x_end_px: int,
                   black_thresh: int = 100, min_dark_pixels: int = 2) -> list[int]:
    """Return list of row indices that contain text (dark pixels)."""
    gray = img.convert("L")
    w, h = gray.size
    rows = []
    pixels = gray.load()
    x_start_px = max(0, x_start_px)
    x_end_px = min(w, x_end_px)
    for y in range(h):
        dark_count = 0
        for x in range(x_start_px, x_end_px):
            if pixels[x, y] < black_thresh:
                dark_count += 1
                if dark_count >= min_dark_pixels:
                    rows.append(y)
                    break
    return rows


def cluster_lines(rows: list[int], gap_threshold: int = 4) -> list[tuple[int, int]]:
    """Cluster contiguous text rows into (top, bottom) bands."""
    if not rows:
        return []
    bands = []
    band_start = rows[0]
    band_end = rows[0]
    for y in rows[1:]:
        if y - band_end <= gap_threshold:
            band_end = y
        else:
            bands.append((band_start, band_end))
            band_start = y
            band_end = y
    bands.append((band_start, band_end))
    return bands


def find_first_glyph_x(img: Image.Image, y_top: int, y_bottom: int,
                       black_thresh: int = 100) -> int:
    """Find the leftmost x with a dark pixel within the band."""
    gray = img.convert("L")
    w, _ = gray.size
    pixels = gray.load()
    for x in range(w):
        for y in range(y_top, min(y_bottom + 1, gray.size[1])):
            if pixels[x, y] < black_thresh:
                return x
    return -1


def main():
    if not WORD_P2.exists():
        print(f"Word PNG missing: {WORD_P2}", file=sys.stderr)
        return 1
    if not OXI_P2.exists():
        print(f"Oxi PNG missing: {OXI_P2}", file=sys.stderr)
        return 1

    for label, path in [("WORD", WORD_P2), ("OXI", OXI_P2)]:
        img = Image.open(path)
        w_px, h_px = img.size
        # Body region in pt: 56.7 to 538.6 (width 481.9)
        # PNG is rendered at some DPI; common is 96, 144, 150, 200
        # Try multiple guesses and report
        for dpi in [96, 150]:
            scale = dpi / 72.0
            body_left_px = int(56.7 * scale)
            body_right_px = int(538.6 * scale)
            if body_right_px <= w_px:
                break
        # If image is 1240px wide and body is 481.9pt, ratio = 1240/page_w_px = scale
        # Estimate actual DPI: image width / (page width in inches)
        # Page width 595.3pt = 595.3/72 in = 8.27 in
        # If image is 1240 wide, dpi = 1240 / 8.27 ≈ 150
        dpi_est = w_px / (595.3 / 72.0)
        scale_est = dpi_est / 72.0
        print(f"\n=== {label} ({w_px}x{h_px}, est dpi={dpi_est:.0f}) ===")
        body_left_px = int(56.7 * scale_est)
        body_right_px = int(538.6 * scale_est)
        print(f"  body in px: left={body_left_px} right={body_right_px}")

        rows = find_text_rows(img, body_left_px, body_right_px)
        bands = cluster_lines(rows)
        print(f"  {len(bands)} text bands found")

        for i, (top, bot) in enumerate(bands[:25]):
            top_pt = top / scale_est
            bot_pt = bot / scale_est
            height_pt = (bot - top + 1) / scale_est
            first_x = find_first_glyph_x(img, top, bot)
            first_x_pt = first_x / scale_est if first_x >= 0 else -1
            indent_pt = first_x_pt - 56.7 if first_x_pt >= 0 else -1
            print(f"  Band {i:2d}: y_top={top_pt:6.1f}pt  y_bot={bot_pt:6.1f}pt  h={height_pt:.1f}  first_x={first_x_pt:.1f}pt  indent={indent_pt:+.1f}pt")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
