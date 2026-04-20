"""Produce 3-column Word/Oxi/diff PNG for a given doc + page.

Usage: python tools/metrics/make_3col_compare.py <doc_stem> <page_num> <out_path>
"""
import sys
from pathlib import Path
import numpy as np
from PIL import Image, ImageDraw, ImageFont

def main(doc_stem, page_num, out_path, crop_box=None):
    repo = Path(__file__).resolve().parents[2]
    word_png = repo / "pipeline_data" / "word_png" / doc_stem / f"page_{page_num:04d}.png"
    oxi_png = repo / "pipeline_data" / "oxi_png" / doc_stem / f"page_{page_num:04d}.png"
    if not word_png.exists() or not oxi_png.exists():
        print(f"Missing: word={word_png.exists()} oxi={oxi_png.exists()}")
        sys.exit(1)

    w = Image.open(word_png).convert("RGB")
    o = Image.open(oxi_png).convert("RGB")

    if crop_box is not None:
        w = w.crop(crop_box)
        o = o.crop(crop_box)

    # Align heights
    H = min(w.height, o.height)
    w = w.crop((0, 0, w.width, H))
    o = o.crop((0, 0, o.width, H))

    # Resize o to match w width if they differ slightly
    if o.width != w.width:
        o = o.resize((w.width, H))

    wa = np.array(w)
    oa = np.array(o)
    diff_mag = np.abs(wa.astype(int) - oa.astype(int)).sum(axis=2)
    overlay = oa.copy()
    mask = diff_mag > 40
    overlay[mask] = [255, 120, 120]
    d = Image.fromarray(overlay.astype(np.uint8))

    gap = 10
    W = w.width * 3 + gap * 2
    canvas = Image.new("RGB", (W, H + 24), "white")
    canvas.paste(w, (0, 24))
    canvas.paste(o, (w.width + gap, 24))
    canvas.paste(d, (w.width * 2 + gap * 2, 24))

    # Header labels
    draw = ImageDraw.Draw(canvas)
    try:
        font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 14)
    except Exception:
        font = None
    draw.text((5, 4), "Word", fill="black", font=font)
    draw.text((w.width + gap + 5, 4), "Oxi", fill="black", font=font)
    draw.text((w.width * 2 + gap * 2 + 5, 4), "Diff (red=mismatch)", fill="red", font=font)

    canvas.save(out_path)
    print(f"Saved {out_path}")

if __name__ == "__main__":
    doc = sys.argv[1]
    page = int(sys.argv[2])
    out = sys.argv[3]
    crop = None
    if len(sys.argv) > 4:
        crop = tuple(int(x) for x in sys.argv[4].split(","))
    main(doc, page, out, crop)
