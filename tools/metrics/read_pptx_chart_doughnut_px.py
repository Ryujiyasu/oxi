# -*- coding: utf-8 -*-
"""Pixel-scan Word's doughnut ring (accent fill), which is more reliable than a
bezier bbox/fit: the control points of a quarter-arc bulge outside the true
circle, so p["rect"] over-states the radius.

For each slide: rasterise the Word PDF, mask the three accent colours, and read
the ring's outer bbox (centre + outer radius) plus the hole radius along the
four cardinal directions from the centre."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import numpy as np
import fitz

PDF = os.path.abspath(
    r"pipeline_data\pptx_probes\chart_doughnut\chart_doughnut.pdf")
DPI = 600.0
S = DPI / 72.0
SX, SY, SW, SH = 72.0, 72.0, 396.0, 288.0
ACCENT = [(79, 129, 189), (192, 80, 77), (155, 187, 89)]

doc = fitz.open(PDF)
for pno in range(doc.page_count):
    pix = doc[pno].get_pixmap(matrix=fitz.Matrix(S / 1.0, S / 1.0))
    a = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, pix.n)[..., :3].astype(int)
    mask = np.zeros(a.shape[:2], bool)
    for c in ACCENT:
        mask |= (np.abs(a - np.array(c)).sum(axis=2) < 40)
    if not mask.any():
        print(f"\n===== slide {pno + 1} ===== (no accent fill)")
        continue
    # The legend swatches carry the same accent colours, so keep only the
    # largest connected blob (the ring itself).
    from scipy import ndimage
    lab, n = ndimage.label(mask)
    if n > 1:
        sizes = ndimage.sum(mask, lab, range(1, n + 1))
        mask = lab == (int(np.argmax(sizes)) + 1)
    ys, xs = np.nonzero(mask)
    x0, x1, y0, y1 = xs.min() / S, xs.max() / S, ys.min() / S, ys.max() / S
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    r_out = ((x1 - x0) + (y1 - y0)) / 4
    print(f"\n===== slide {pno + 1} =====")
    print(f"  ring bbox x[{x0:7.2f},{x1:7.2f}] y[{y0:7.2f},{y1:7.2f}]"
          f"  w={x1-x0:6.2f} h={y1-y0:6.2f}")
    print(f"  centre=({cx:.2f},{cy:.2f}) = (sx+{cx-SX:.2f}, sy+{cy-SY:.2f})"
          f"   frame_cx={SX+SW/2:.2f}")
    print(f"  top={y0:.2f} (sy+{y0-SY:.2f})  bot={y1:.2f} (sy+shh{y1-(SY+SH):+.2f})"
          f"  r_out={r_out:.2f}")

    # hole radius: walk outward from the centre until accent starts
    holes = []
    icx, icy = int(round(cx * S)), int(round(cy * S))
    for dx, dy, lbl in ((1, 0, "E"), (-1, 0, "W"), (0, -1, "N"), (0, 1, "S")):
        for step in range(1, int(r_out * S) + 1):
            px, py = icx + dx * step, icy + dy * step
            if 0 <= py < mask.shape[0] and 0 <= px < mask.shape[1] and mask[py, px]:
                holes.append((lbl, step / S))
                break
    if holes:
        vals = [v for _, v in holes]
        r_in = sum(vals) / len(vals)
        print("  hole edge: " + "  ".join(f"{l}={v:.2f}" for l, v in holes))
        print(f"  r_in={r_in:.2f}   r_in/r_out={r_in/r_out:.4f}")
