"""Pagination × IoU cross-axis classification.

R84 (2026-04-29): companion to R83's iou_ssim_gap.py. Where R83 cross-cut
Phase 2 (IoU) and Phase 3 (SSIM) to find rendering-only candidates, R84
cross-cuts Phase 1 (pagination) and Phase 2 (IoU) to classify Phase 1
FAIL docs and find Phase 2 priority candidates.

Four classifications:

| Phase 1   | IoU      | Class                    | Implication                      |
|-----------|----------|--------------------------|----------------------------------|
| FAIL      | >= 0.95  | "pure pagination bug"    | surgical page-break fix candidate|
| FAIL      | 0.7-0.95 | "mixed"                  | both page-break + position drift |
| FAIL      | < 0.7    | "deep cascade"           | R44 multi-week wall              |
| PASS      | < 0.99   | "Phase 2 priority"       | within-page position needs work  |
| PASS      | >= 0.99  | "both pass" (omitted)    | Phase 3 (SSIM) is the only gate  |

Output:
    pipeline_data/pagination_iou_class/_summary.json
    stdout                                          human-readable per class

Run:
    python tools/metrics/pagination_iou_class.py
"""
from __future__ import annotations

import json
import os
import sys

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
PAG_PATH = os.path.join(REPO_ROOT, "pipeline_data", "pagination_diff", "_summary.json")
IOU_PATH = os.path.join(REPO_ROOT, "pipeline_data", "element_iou_diff", "_summary.json")
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_iou_class")
OUT_PATH = os.path.join(OUT_DIR, "_summary.json")

IOU_PASS = 0.99
IOU_NEAR_PASS = 0.95
IOU_DEEP_CASCADE = 0.7


def classify(p1_pass: bool, iou: float) -> str:
    if p1_pass:
        return "phase2_priority" if iou < IOU_PASS else "both_pass"
    # Phase 1 FAIL classes
    if iou >= IOU_NEAR_PASS:
        return "pure_pagination_bug"
    if iou >= IOU_DEEP_CASCADE:
        return "mixed"
    return "deep_cascade"


def main() -> int:
    if not os.path.exists(PAG_PATH):
        print(f"missing: {PAG_PATH}", file=sys.stderr)
        return 1
    if not os.path.exists(IOU_PATH):
        print(f"missing: {IOU_PATH}", file=sys.stderr)
        return 1

    with open(PAG_PATH, encoding="utf-8") as f:
        pag = json.load(f)
    with open(IOU_PATH, encoding="utf-8") as f:
        iou = json.load(f)

    iou_per_doc = {d["doc_id"]: d["mean_iou"] for d in iou.get("docs", [])}

    classes: dict[str, list[dict]] = {
        "pure_pagination_bug": [],
        "mixed": [],
        "deep_cascade": [],
        "phase2_priority": [],
        "both_pass": [],
    }

    for d in pag.get("docs", []):
        doc_id = d["doc_id"]
        p1_pass = bool(d.get("pass"))
        p1_score = float(d.get("score", 0.0))
        iou_val = iou_per_doc.get(doc_id)
        if iou_val is None:
            continue
        cls = classify(p1_pass, iou_val)
        classes[cls].append({
            "doc_id": doc_id,
            "pagination_score": p1_score,
            "mean_iou": iou_val,
        })

    # Sort each class meaningfully:
    # FAIL classes: by IoU descending (closer-to-pass first)
    # PASS classes: by IoU ascending (worst first = priority)
    classes["pure_pagination_bug"].sort(key=lambda r: -r["mean_iou"])
    classes["mixed"].sort(key=lambda r: -r["mean_iou"])
    classes["deep_cascade"].sort(key=lambda r: -r["mean_iou"])
    classes["phase2_priority"].sort(key=lambda r: r["mean_iou"])
    classes["both_pass"].sort(key=lambda r: r["doc_id"])

    # Print human-readable summary.
    print(f"Pagination x IoU classification (post-R83 baseline)\n")
    print(f"  Class                  | Count | Strategy")
    print(f"  -----------------------+-------+--------------------------")
    print(f"  pure_pagination_bug    | {len(classes['pure_pagination_bug']):5d} | surgical page-break fix")
    print(f"  mixed                  | {len(classes['mixed']):5d} | both pag + position drift")
    print(f"  deep_cascade           | {len(classes['deep_cascade']):5d} | R44 multi-week wall")
    print(f"  phase2_priority        | {len(classes['phase2_priority']):5d} | within-page position work")
    print(f"  both_pass              | {len(classes['both_pass']):5d} | Phase 3 (SSIM) is the gate")
    total = sum(len(v) for v in classes.values())
    print(f"  TOTAL                  | {total:5d} |")

    # Top candidates per class.
    def show(class_name: str, n: int = 5):
        rows = classes[class_name]
        if not rows:
            return
        print(f"\n  Top {min(n, len(rows))} in '{class_name}':")
        for r in rows[:n]:
            print(f"    {r['doc_id']:45s}  pag={r['pagination_score']:.4f}  IoU={r['mean_iou']:.4f}")

    show("pure_pagination_bug")
    show("mixed")
    show("phase2_priority", n=8)

    # Persist.
    os.makedirs(OUT_DIR, exist_ok=True)
    summary = {
        "thresholds": {
            "iou_pass": IOU_PASS,
            "iou_near_pass": IOU_NEAR_PASS,
            "iou_deep_cascade": IOU_DEEP_CASCADE,
        },
        "counts": {k: len(v) for k, v in classes.items()},
        "classes": classes,
    }
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
