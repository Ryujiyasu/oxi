# -*- coding: utf-8 -*-
"""S558 — targeted SSIM check for the s475_pair split fix. Renders the listed
doc_ids via the active Oxi renderer (DWrite default), computes per-page SSIM
vs the cached word_png, and compares to ssim_baseline.json. Reports
improvements/regressions per the Phase-1 sentinel (mean drop must be small).
Caller MUST clear pipeline_data/oxi_png/<doc_id> first (render_with_oxi skips
cached PNGs).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from pipeline.config import OXI_PNG_DIR, WORD_PNG_DIR
from pipeline.oxi_renderer import render_with_oxi
from pipeline.ssim_calculator import calculate_ssim
from pipeline.verify import load_baseline
from pathlib import Path

DOCS_DIR = os.path.join(os.path.dirname(__file__), "..", "..",
                        "tools", "golden-test", "documents", "docx")

targets = [l.strip() for l in open(sys.argv[1]) if l.strip()] if len(sys.argv) > 1 else []
# map doc_id (baseline key) -> docx path
docx_paths = []
for t in targets:
    p = os.path.join(DOCS_DIR, t + ".docx")
    if os.path.exists(p):
        docx_paths.append(os.path.abspath(p))
    else:
        print("missing docx:", t)

baseline = load_baseline()
oxi = render_with_oxi(docx_paths)
# build word_results from cached word_png
word = {}
for dp in docx_paths:
    doc_id = Path(dp).stem
    wd = Path(WORD_PNG_DIR) / doc_id
    if wd.exists():
        word[dp] = sorted(str(p) for p in wd.glob("page_*.png"))

scores = calculate_ssim(word, oxi, skip_heatmap=True)

sys.stdout.reconfigure(encoding="utf-8")
imp = reg = 0
gain = loss = 0.0
per_doc = {}
for s in scores:
    doc_id = s["doc_id"]; page = str(s["page"]); new = s["ssim_score"]
    pk = page if (doc_id in baseline and page in baseline.get(doc_id, {})) else f"{int(page):04d}"
    if doc_id in baseline and pk in baseline[doc_id]:
        old = baseline[doc_id][pk]
        diff = new - old
        per_doc.setdefault(doc_id, []).append(diff)
        if diff < -0.001:
            reg += 1; loss += -diff
            print(f"  REGRESS {doc_id} p{page}: {old:.4f} -> {new:.4f} ({diff:+.4f})")
        elif diff > 0.001:
            imp += 1; gain += diff
            print(f"  improve {doc_id} p{page}: {old:.4f} -> {new:.4f} ({diff:+.4f})")
print("\n=== per-doc mean delta ===")
for d, ds in sorted(per_doc.items()):
    print(f"  {d}: meanΔ={sum(ds)/len(ds):+.4f}  ({len(ds)} pages)")
print(f"\nTOTAL improved={imp} regressed={reg}  gain={gain:.4f} loss={loss:.4f} net={gain-loss:+.4f}")
