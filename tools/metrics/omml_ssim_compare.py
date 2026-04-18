"""Compute SSIM between Oxi PNG and Word PNG for each OMML fixture.

Outputs a table showing how close Oxi's rendering is to Word's.
Identifies the best/worst fixtures for future refinement.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from pipeline.ssim_calculator import calculate_ssim

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OXI_DIR = Path(__file__).resolve().parent.parent.parent / "pipeline_data" / "oxi_omml"
WORD_DIR = Path(__file__).resolve().parent.parent.parent / "pipeline_data" / "word_omml"

# Build fake "renderer_results" dicts for calculate_ssim
fixtures = sorted([p.stem for p in (Path(__file__).resolve().parent.parent / "fixtures" / "omml_samples").glob("*.docx")])

word_res = {}
oxi_res = {}
for name in fixtures:
    word_png = WORD_DIR / name / "page_0001.png"
    oxi_png_alt1 = OXI_DIR / f"{name}_p1.png"
    oxi_png_alt2 = OXI_DIR / f"{name}" / "page_0001.png"
    if not word_png.exists():
        print(f"skip {name}: word png missing")
        continue
    if oxi_png_alt1.exists():
        oxi_png = oxi_png_alt1
    elif oxi_png_alt2.exists():
        oxi_png = oxi_png_alt2
    else:
        print(f"skip {name}: oxi png missing (tried {oxi_png_alt1}, {oxi_png_alt2})")
        continue
    docx_key = f"tools/fixtures/omml_samples/{name}.docx"
    word_res[docx_key] = [str(word_png)]
    oxi_res[docx_key] = [str(oxi_png)]

print(f"Comparing {len(word_res)} fixtures...")
scores = calculate_ssim(word_res, oxi_res, skip_heatmap=True)

# Summary table
print(f"\n{'fixture':<25} {'SSIM':>8}")
print("-" * 36)
total = 0.0
n = 0
for s in scores:
    doc_id = s["doc_id"]
    ssim = s["ssim_score"]
    total += ssim
    n += 1
    mark = "✓" if ssim >= 0.8 else ("~" if ssim >= 0.6 else "✗")
    print(f"{mark} {doc_id[:25]:<25} {ssim:>8.4f}")

if n > 0:
    print(f"\nMean SSIM: {total/n:.4f} across {n} fixtures")
