"""IoU vs SSIM gap diagnostic — find rendering-only improvement candidates.

R83 (2026-04-29): formalises the ad-hoc analysis that surfaced after R79's
dual-axis sentinel validation. Cross-references Phase 2 (`pipeline_data/element_iou_diff/_summary.json`,
position-IoU per doc) and Phase 3 (`pipeline_data/ssim_baseline.json`, per-page
SSIM) to answer:

    "Which docs pass position-IoU but still have low SSIM?"

A doc with high IoU + low SSIM means paragraph positions match Word's
output (Phase 2 happy) but pixel-level rendering still diverges (Phase 3
gap). The gap is "rendering-only" — character spacing, font metrics,
border thickness, glyph kerning, etc — and would NOT be surfaced by
either Phase 1 (pagination) or Phase 2 (IoU) gates alone.

Output:
    pipeline_data/iou_ssim_gap/_summary.json   structured per-doc data
    stdout                                     human-readable table

Run:
    python tools/metrics/iou_ssim_gap.py
"""
from __future__ import annotations

import json
import os
import sys

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
IOU_PATH = os.path.join(REPO_ROOT, "pipeline_data", "element_iou_diff", "_summary.json")
SSIM_PATH = os.path.join(REPO_ROOT, "pipeline_data", "ssim_baseline.json")
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "iou_ssim_gap")
OUT_PATH = os.path.join(OUT_DIR, "_summary.json")

# Threshold for "passing IoU" — matches Phase 2 gate.
IOU_PASS = 0.99


def load_json(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def main() -> int:
    if not os.path.exists(IOU_PATH):
        print(f"missing: {IOU_PATH}", file=sys.stderr)
        return 1
    if not os.path.exists(SSIM_PATH):
        print(f"missing: {SSIM_PATH}", file=sys.stderr)
        return 1

    iou = load_json(IOU_PATH)
    ssim_data = load_json(SSIM_PATH)

    # Per-doc IoU lookup (key = doc_id)
    iou_per_doc = {d["doc_id"]: d["mean_iou"] for d in iou.get("docs", [])}

    # Per-doc SSIM mean — keys in ssim_baseline are full filenames; reduce
    # to short doc_id by taking the leading 12-char hex prefix (matches the
    # convention used by element_iou_diff).
    ssim_per_doc: dict[str, float] = {}
    for full_name, pages in ssim_data.items():
        if not pages:
            continue
        prefix = full_name.split("_")[0]
        vals = list(pages.values())
        ssim_per_doc[prefix] = sum(vals) / len(vals)

    # Build the cross-join. For each IoU doc with mean >= IOU_PASS, find
    # the matching SSIM entry by prefix.
    results: list[dict] = []
    for doc_id, doc_iou in iou_per_doc.items():
        if doc_iou < IOU_PASS:
            continue
        ssim = ssim_per_doc.get(doc_id[:12])
        if ssim is None:
            # Try shorter prefixes (some doc_ids don't preserve hex)
            for full, val in ssim_per_doc.items():
                if doc_id.startswith(full[:6]) or full.startswith(doc_id[:6]):
                    ssim = val
                    break
        results.append({
            "doc_id": doc_id,
            "mean_iou": doc_iou,
            "mean_ssim": ssim,
            "gap": (doc_iou - ssim) if ssim is not None else None,
        })

    # Sort by gap descending (biggest "rendering-only" candidates first).
    # None-gap entries go to the end.
    results.sort(key=lambda r: (-(r["gap"] if r["gap"] is not None else -1.0)))

    # Print human-readable table.
    print(f"Phase 2 IoU pass threshold: {IOU_PASS}")
    print(f"Docs passing IoU but with measurable SSIM gap (top 20):\n")
    print(f"  {'doc_id':45s}  {'IoU':>6s}  {'SSIM':>6s}  {'gap':>7s}")
    print("  " + "-" * 65)
    for r in results[:20]:
        if r["mean_ssim"] is None:
            print(f"  {r['doc_id']:45s}  {r['mean_iou']:.4f}  {'?':>6s}  {'?':>7s}")
        else:
            print(
                f"  {r['doc_id']:45s}  {r['mean_iou']:.4f}  {r['mean_ssim']:.4f}  {r['gap']:+.4f}"
            )

    # Histogram-ish summary.
    bands = {">=0.20": 0, "0.10-0.20": 0, "0.05-0.10": 0, "0.02-0.05": 0, "<0.02": 0}
    for r in results:
        if r["gap"] is None:
            continue
        if r["gap"] >= 0.20:
            bands[">=0.20"] += 1
        elif r["gap"] >= 0.10:
            bands["0.10-0.20"] += 1
        elif r["gap"] >= 0.05:
            bands["0.05-0.10"] += 1
        elif r["gap"] >= 0.02:
            bands["0.02-0.05"] += 1
        else:
            bands["<0.02"] += 1
    print("\nGap distribution (rendering-only candidate severity):")
    for band, count in bands.items():
        print(f"  {band:>10s} : {count:3d} docs")

    # Persist.
    os.makedirs(OUT_DIR, exist_ok=True)
    summary = {
        "iou_pass_threshold": IOU_PASS,
        "n_total_iou_pass": len(results),
        "n_with_ssim_match": sum(1 for r in results if r["mean_ssim"] is not None),
        "gap_bands": bands,
        "docs": results,
    }
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
