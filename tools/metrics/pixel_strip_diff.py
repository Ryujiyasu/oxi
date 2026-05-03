"""Compute Word vs Oxi pixel diff per horizontal strip (20pt bands).
Identifies which Y region has biggest pixel mismatch — highest-leverage
investigation target."""
import sys
from PIL import Image
import numpy as np

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WORD_PNG = "pipeline_data/word_png/b35123fe8efc_tokumei_08_01/page_0001.png"
OXI_PNG = "pipeline_data/oxi_png/b35123fe8efc_tokumei_08_01/page_p1.png"

word = np.array(Image.open(WORD_PNG).convert("L"))
oxi = np.array(Image.open(OXI_PNG).convert("L"))

# Resize Oxi to Word dimensions if needed
if word.shape != oxi.shape:
    from PIL import Image as PI
    oxi_pil = PI.open(OXI_PNG).convert("L").resize((word.shape[1], word.shape[0]))
    oxi = np.array(oxi_pil)

h, w = word.shape
page_h_pt = 841.9
page_w_pt = 595.30

print(f"Image: {w}x{h}, page {page_w_pt}x{page_h_pt}pt")
print(f"\n{'y range':>16} {'mean_diff':>10} {'pct>50':>8} {'note':>30}")
print("-" * 72)

strip_pt = 20
strips = []
for y_pt_start in range(0, int(page_h_pt), strip_pt):
    y_px_start = int(y_pt_start * h / page_h_pt)
    y_px_end = int((y_pt_start + strip_pt) * h / page_h_pt)
    strip_w = word[y_px_start:y_px_end, :]
    strip_o = oxi[y_px_start:y_px_end, :]
    diff = np.abs(strip_w.astype(np.int32) - strip_o.astype(np.int32))
    mean_diff = diff.mean()
    pct_big_diff = (diff > 50).sum() / diff.size * 100
    strips.append((y_pt_start, mean_diff, pct_big_diff))
    print(f"{y_pt_start:>4}-{y_pt_start+strip_pt:<4}pt   {mean_diff:>10.2f} {pct_big_diff:>7.1f}%")

print("\n=== Top 5 worst strips ===")
strips.sort(key=lambda s: -s[1])
for y_start, mean_d, pct in strips[:5]:
    print(f"y={y_start}-{y_start+strip_pt}pt: mean={mean_d:.2f} >50%={pct:.1f}%")
