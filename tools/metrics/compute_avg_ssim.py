"""Compute current avg SSIM from cached oxi_png + word_png.

Re-runs SSIM calc on existing PNGs (no re-render). Reports total avg.
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from pipeline.ssim_calculator import calculate_ssim
from pipeline.config import OXI_PNG_DIR, WORD_PNG_DIR
from pathlib import Path

# Build doc_id -> [(word_png, oxi_png), ...] mapping
word_results = {}
oxi_results = {}
for d in sorted(Path(OXI_PNG_DIR).iterdir()):
    if not d.is_dir():
        continue
    doc_id = d.name
    oxi_pngs = sorted(d.glob("page_*.png"))
    word_dir = Path(WORD_PNG_DIR) / doc_id
    if not word_dir.exists():
        continue
    word_pngs = sorted(word_dir.glob("page_*.png"))
    if oxi_pngs and word_pngs:
        # Use docx_path key as expected by calculate_ssim
        docx_path = doc_id  # any unique key works
        word_results[docx_path] = [str(p) for p in word_pngs]
        oxi_results[docx_path] = [str(p) for p in oxi_pngs]

scores = calculate_ssim(word_results, oxi_results, skip_heatmap=True)
if not scores:
    print("No scores computed")
    sys.exit(1)

avg = sum(s["ssim_score"] for s in scores) / len(scores)
print(f"Avg SSIM: {avg:.4f}")
print(f"Pages: {len(scores)}")
print(f">0.95: {sum(1 for s in scores if s['ssim_score'] > 0.95)}")
print(f">0.90: {sum(1 for s in scores if s['ssim_score'] > 0.90)}")
print(f"<0.50: {sum(1 for s in scores if s['ssim_score'] < 0.50)}")
