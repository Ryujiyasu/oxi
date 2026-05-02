"""Compute Word vs Oxi pixel diff per horizontal strip for d77a p.7.

d77a p.7 SSIM = 0.6268 (genuine bug — same in old and new baseline).
Identifies which Y region is the biggest pixel mismatch."""
import sys
from PIL import Image
import numpy as np

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WORD_PNG = "pipeline_data/word_png/d77a58485f16_20240705_resources_data_outline_08/page_0007.png"
OXI_PNG = "pipeline_data/oxi_png/d77a58485f16_20240705_resources_data_outline_08/page_0007.png"

word = np.array(Image.open(WORD_PNG).convert("L"))
oxi = np.array(Image.open(OXI_PNG).convert("L"))

if word.shape != oxi.shape:
    from PIL import Image as PI
    oxi_pil = PI.open(OXI_PNG).convert("L").resize((word.shape[1], word.shape[0]))
    oxi = np.array(oxi_pil)

h, w = word.shape
page_h_pt = 841.9
page_w_pt = 595.30

print(f"Image: {w}x{h}, page {page_w_pt}x{page_h_pt}pt")
print(f"{'y range (pt)':>16} {'mean_diff':>10} {'pct>50':>8} {'pct>100':>8}")
print("-" * 50)

strip_pt = 20
strips = []
for y_pt_start in range(0, int(page_h_pt), strip_pt):
    y_pt_end = min(y_pt_start + strip_pt, page_h_pt)
    y_px_start = int(y_pt_start * h / page_h_pt)
    y_px_end = int(y_pt_end * h / page_h_pt)
    if y_px_start >= y_px_end:
        continue
    w_strip = word[y_px_start:y_px_end, :]
    o_strip = oxi[y_px_start:y_px_end, :]
    diff = np.abs(w_strip.astype(int) - o_strip.astype(int))
    mean_d = float(diff.mean())
    pct50 = 100.0 * (diff > 50).sum() / diff.size
    pct100 = 100.0 * (diff > 100).sum() / diff.size
    strips.append((y_pt_start, y_pt_end, mean_d, pct50, pct100))
    print(f"{y_pt_start:5d}-{y_pt_end:5.0f}    {mean_d:10.2f} {pct50:7.1f}% {pct100:7.1f}%")

print(f"\nWorst strips by mean_diff:")
strips.sort(key=lambda s: -s[2])
for y0, y1, md, p50, p100 in strips[:5]:
    print(f"  y={y0:.0f}-{y1:.0f}pt  mean={md:.2f}  pct>50={p50:.1f}%  pct>100={p100:.1f}%")
