"""Detect table row boundaries in Word/Oxi PNG by horizontal-line scanning.

Find Y positions where horizontal dark lines span > N% of page width.
These are typically table cell borders. Compare Word vs Oxi row Y values
to identify per-row vertical drift.
"""
import sys
import os
from PIL import Image
import numpy as np

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WORD_PNG = "pipeline_data/word_png/b35123fe8efc_tokumei_08_01/page_0001.png"
OXI_PNG = "pipeline_data/oxi_png/b35123fe8efc_tokumei_08_01/page_p1.png"


def detect_horizontal_lines(png_path, min_span_pct=0.20, dark_thresh=80):
    """Return list of y_pt values where a horizontal dark line spans
    more than min_span_pct of the page width."""
    img = Image.open(png_path).convert("L")
    arr = np.array(img)
    h, w = arr.shape
    page_w_pt = 595.30
    page_h_pt = 841.90
    # For each y row, count consecutive dark pixels in horizontal stretches
    lines = []
    for y_px in range(h):
        row = arr[y_px, :]
        dark = (row < dark_thresh).astype(int)
        # Find longest contiguous dark run
        max_run = 0; cur_run = 0
        for v in dark:
            if v: cur_run += 1; max_run = max(max_run, cur_run)
            else: cur_run = 0
        if max_run / w > min_span_pct:
            lines.append((y_px, max_run / w))
    # Cluster adjacent y_px into single line
    clusters = []
    if lines:
        cur_start = lines[0][0]; cur_end = lines[0][0]; cur_max_pct = lines[0][1]
        for y_px, pct in lines[1:]:
            if y_px - cur_end <= 2:
                cur_end = y_px
                cur_max_pct = max(cur_max_pct, pct)
            else:
                clusters.append((cur_start, cur_end, cur_max_pct))
                cur_start = y_px; cur_end = y_px; cur_max_pct = pct
        clusters.append((cur_start, cur_end, cur_max_pct))
    # Convert px → pt
    return [(s * page_h_pt / h, e * page_h_pt / h, p) for s, e, p in clusters]


def main():
    print("=== Word PNG ===", flush=True)
    word_lines = detect_horizontal_lines(WORD_PNG)
    print(f"Found {len(word_lines)} horizontal lines:", flush=True)
    for s, e, p in word_lines:
        print(f"  y_pt={s:.1f}-{e:.1f} span={p:.0%}", flush=True)

    print("\n=== Oxi PNG ===", flush=True)
    oxi_lines = detect_horizontal_lines(OXI_PNG)
    print(f"Found {len(oxi_lines)} horizontal lines:", flush=True)
    for s, e, p in oxi_lines:
        print(f"  y_pt={s:.1f}-{e:.1f} span={p:.0%}", flush=True)

    # Pair Word vs Oxi lines (closest y match)
    print("\n=== Word vs Oxi row pairing ===", flush=True)
    print(f"{'word_y':>8} {'oxi_y':>8} {'diff':>8} {'span_w':>6} {'span_o':>6}",
          flush=True)
    used_oxi = set()
    for w_s, _, w_pct in word_lines:
        # Find closest Oxi line not yet used
        best = None; best_dist = float("inf")
        for i, (o_s, _, o_pct) in enumerate(oxi_lines):
            if i in used_oxi: continue
            d = abs(o_s - w_s)
            if d < best_dist:
                best_dist = d; best = (i, o_s, o_pct)
        if best and best_dist < 30:
            used_oxi.add(best[0])
            diff = best[1] - w_s
            marker = " ⚠" if abs(diff) > 5 else ""
            print(f"{w_s:>8.1f} {best[1]:>8.1f} {diff:>+8.2f} {w_pct:>6.0%} {best[2]:>6.0%}{marker}",
                  flush=True)
        else:
            print(f"{w_s:>8.1f}  (NO MATCH)", flush=True)


if __name__ == "__main__":
    main()
