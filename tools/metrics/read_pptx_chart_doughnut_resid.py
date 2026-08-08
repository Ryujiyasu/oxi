# -*- coding: utf-8 -*-
"""Read the doughnut-residual probe: ring geometry by 600dpi pixel scan,
legend swatches/labels from the vector+text layers."""
import sys, os

sys.stdout.reconfigure(encoding="utf-8")
import fitz
import numpy as np
from scipy import ndimage

base = os.path.abspath(r"pipeline_data\pptx_probes\chart_doughnut_resid")
pdf = os.path.join(base, "chart_doughnut_resid.pdf")
DPI = 600
S = DPI / 72.0
ACC = [(0x4F, 0x81, 0xBD), (0xC0, 0x50, 0x4D), (0x9B, 0xBB, 0x59)]

doc = fitz.open(pdf)
print(f"pages={doc.page_count}  frame sx=72 sy=72 sw=396 shh=288\n")

for pno in range(doc.page_count):
    page = doc[pno]
    pix = page.get_pixmap(dpi=DPI)
    a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)[
        :, :, :3
    ]
    print(f"=== S{pno+1} ===")

    # --- ring geometry: keep only the LARGEST connected blob of each accent
    #     colour (legend swatches share the colour and skew the bbox).
    mask = np.zeros(a.shape[:2], bool)
    for c in ACC:
        d = np.abs(a.astype(int) - np.array(c)).sum(2)
        m = d < 40
        if m.sum() < 100:
            continue
        lab, n = ndimage.label(m)
        if n:
            sizes = ndimage.sum(m, lab, range(1, n + 1))
            mask |= lab == (int(np.argmax(sizes)) + 1)
    if mask.sum() > 1000:
        ys, xs = np.where(mask)
        x0, x1, y0, y1 = xs.min() / S, xs.max() / S, ys.min() / S, ys.max() / S
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        r = ((x1 - x0) + (y1 - y0)) / 4
        print(
            f"  ring  x {x0:7.2f}..{x1:7.2f}  y {y0:7.2f}..{y1:7.2f} "
            f" c=({cx:.2f},{cy:.2f}) r={r:.2f}  w={x1-x0:.2f} h={y1-y0:.2f}"
        )
        print(
            f"        top-sy={y0-72:6.2f}  bot-(sy+shh)={y1-360:6.2f} "
            f" cx-frame_c={cx-270:6.2f}"
        )
        # hole radius along the 4 cardinals from the blob centre
        ccx, ccy = int(cx * S), int(cy * S)
        holes = []
        for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            k = 0
            while k < int(r * S) and not mask[ccy + dy * k, ccx + dx * k]:
                k += 1
            holes.append(k / S)
        print(
            "        hole r: "
            + " ".join(f"{h:.2f}" for h in holes)
            + f"   ratio={np.mean(holes)/r:.4f}"
        )
        # concentric rings?  count colour bands along +x
        row = mask[ccy, ccx:]
        edges, prev = [], False
        for i, v in enumerate(row):
            if v != prev:
                edges.append(i / S)
                prev = v
        print("        +x band edges (pt from centre): " + " ".join(f"{e:.2f}" for e in edges[:8]))

    # --- legend swatches (small accent rects) + label spans
    for d in page.get_drawings():
        f = d.get("fill")
        if not f:
            continue
        rgb = tuple(int(round(v * 255)) for v in f)
        if any(abs(rgb[0] - c[0]) + abs(rgb[1] - c[1]) + abs(rgb[2] - c[2]) < 40 for c in ACC):
            r_ = d["rect"]
            if r_.width < 30 and r_.height < 30:
                print(
                    f"  swatch {rgb}  x {r_.x0:7.2f}..{r_.x1:7.2f} "
                    f" y {r_.y0:7.2f}..{r_.y1:7.2f}  w={r_.width:.2f} h={r_.height:.2f}"
                )
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                t = "".join(c["c"] for c in s["chars"])
                if t.strip():
                    print(
                        f"  span  ({s['origin'][0]:7.2f},{s['origin'][1]:7.2f}) "
                        f"x1={s['bbox'][2]:7.2f} sz={s['size']:5.2f} {s['font'][:22]:22s} {t!r}"
                    )
    print()
