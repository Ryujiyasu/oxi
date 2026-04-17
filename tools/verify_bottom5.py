"""Verify only bottom-5 docs against baseline — skip the 177-doc nuclear option.

Usage: python tools/verify_bottom5.py
"""
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from pipeline.word_renderer import render_with_word
from pipeline.oxi_renderer import render_with_oxi
from pipeline.ssim_calculator import calculate_ssim
from pipeline.baseline import load_baseline

DOCX_DIR = Path(__file__).parent.parent / "tools/golden-test/documents/docx"
BOTTOM5 = [
    "683ffcab86e2_20230331_resources_open_data_contract_addon_00",
    "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00",
    "d77a58485f16_20240705_resources_data_outline_08",
    "b35123fe8efc_tokumei_08_01",
    "b837808d0555_20240705_resources_data_guideline_02",
]

baseline = load_baseline()
paths = [str(DOCX_DIR / f"{d}.docx") for d in BOTTOM5]
paths = [p for p in paths if os.path.exists(p)]
print(f"Verifying {len(paths)} bottom-5 docs...")

word_results = render_with_word(paths)
oxi_results = render_with_oxi(paths)
scores = calculate_ssim(word_results, oxi_results, skip_heatmap=True)

print(f"\n{'doc':<55} {'page':>4} {'old':>7} {'new':>7} {'diff':>+8}")
print("-" * 90)
doc_min_new = {}
doc_min_old = {}
for s in scores:
    doc = s["doc_id"]; p = str(s["page"]); new = s["ssim_score"]
    old = baseline.get(doc, {}).get(p, baseline.get(doc, {}).get(f"{int(p):04d}"))
    if old is None:
        print(f"  {doc[:55]:<55} {p:>4}    n/a  {new:.4f}")
        continue
    diff = new - old
    mark = " " if abs(diff) < 0.001 else ("!" if diff < -0.001 else "+")
    print(f"{mark} {doc[:55]:<55} {p:>4}  {old:.4f}  {new:.4f}  {diff:+.4f}")
    # track doc-min
    doc_min_new[doc] = min(doc_min_new.get(doc, 1.0), new)
    doc_min_old[doc] = min(doc_min_old.get(doc, 1.0), old)

print("\nPer-doc min (worst page per doc):")
for d in BOTTOM5:
    on = doc_min_old.get(d); nn = doc_min_new.get(d)
    if on is None or nn is None: continue
    print(f"  {d[:60]:<60}  {on:.4f} -> {nn:.4f}  Δ{nn-on:+.4f}")

pre5 = sum(doc_min_old[d] for d in BOTTOM5 if d in doc_min_old)
post5 = sum(doc_min_new[d] for d in BOTTOM5 if d in doc_min_new)
print(f"\nbottom-5 floor sum: {pre5:.4f} -> {post5:.4f}  Δ{post5-pre5:+.4f}")
print(f"merge gate (post > pre, Phase 2): {'PASS' if post5 > pre5 else 'FAIL'}")
