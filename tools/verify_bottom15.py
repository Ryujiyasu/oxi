"""Verify bottom-15 worst docs — broader than bottom-5 for confidence merges.

Usage: python tools/verify_bottom15.py
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

baseline = load_baseline()
DOCX_DIR = Path(__file__).parent.parent / "tools/golden-test/documents/docx"

# Compute bottom-15 from baseline
mins = []
for d, pgs in baseline.items():
    if pgs:
        mn = min(pgs.values())
        mins.append((mn, d))
mins.sort()
bottom15 = [d for (_, d) in mins[:15]]
print(f"Bottom-15: {bottom15}")

paths = [str(DOCX_DIR / f"{d}.docx") for d in bottom15]
paths = [p for p in paths if os.path.exists(p)]
print(f"Verifying {len(paths)} docs...")

word_results = render_with_word(paths)
oxi_results = render_with_oxi(paths)
scores = calculate_ssim(word_results, oxi_results, skip_heatmap=True)

doc_min_new = {}
doc_min_old = {}
regressions = []
for s in scores:
    doc = s["doc_id"]; p = str(s["page"]); new = s["ssim_score"]
    old = baseline.get(doc, {}).get(p, baseline.get(doc, {}).get(f"{int(p):04d}"))
    if old is None:
        continue
    diff = new - old
    if diff < -0.001:
        regressions.append((doc, p, old, new, diff))
    doc_min_new[doc] = min(doc_min_new.get(doc, 1.0), new)
    doc_min_old[doc] = min(doc_min_old.get(doc, 1.0), old)

print("\nPer-doc min:")
for d in bottom15:
    on = doc_min_old.get(d); nn = doc_min_new.get(d)
    if on is None or nn is None: continue
    mark = "!" if nn < on - 0.001 else ("+" if nn > on + 0.001 else " ")
    print(f"  {mark} {d[:55]:<55}  {on:.4f} -> {nn:.4f}  d={nn-on:+.4f}")

if regressions:
    print(f"\n{len(regressions)} page regressions:")
    for doc, p, old, new, diff in regressions[:15]:
        print(f"  {doc[:50]:<50} p.{p}  {old:.4f} -> {new:.4f}  d={diff:+.4f}")

pre = sum(doc_min_old[d] for d in bottom15 if d in doc_min_old)
post = sum(doc_min_new[d] for d in bottom15 if d in doc_min_new)
print(f"\nbottom-15 floor sum: {pre:.4f} -> {post:.4f}  d={post-pre:+.4f}")
