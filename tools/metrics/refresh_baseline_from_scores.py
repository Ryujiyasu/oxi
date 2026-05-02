"""
Refresh pipeline_data/ssim_baseline.json with values from a verify-run
ssim_scores JSON. Only updates pages that exist in the scores file —
unaffected docs/pages are untouched.

Usage:
  python tools/metrics/refresh_baseline_from_scores.py <scores.json>
"""
import json
import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
BASELINE = os.path.join(ROOT, "pipeline_data", "ssim_baseline.json")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    scores_path = sys.argv[1]
    with open(scores_path, "r", encoding="utf-8") as f:
        scores = json.load(f)
    with open(BASELINE, "r", encoding="utf-8") as f:
        baseline = json.load(f)

    changes = []
    for s in scores:
        doc_id = s["doc_id"]
        page = str(s["page"])
        new = float(s["ssim_score"])
        old = None
        if doc_id in baseline:
            # baseline keys may be either "1" or "0001" — try both
            if page in baseline[doc_id]:
                key = page
            elif f"{int(page):04d}" in baseline[doc_id]:
                key = f"{int(page):04d}"
            else:
                key = page
                baseline[doc_id][key] = new
                changes.append((doc_id, page, None, new))
                continue
            old = float(baseline[doc_id][key])
        else:
            baseline[doc_id] = {}
            key = page
            baseline[doc_id][key] = new
            changes.append((doc_id, page, None, new))
            continue

        if abs(new - old) > 1e-6:
            baseline[doc_id][key] = new
            changes.append((doc_id, page, old, new))

    with open(BASELINE, "w", encoding="utf-8") as f:
        json.dump(baseline, f, indent=2)

    print(f"# refreshed {BASELINE}")
    print(f"# {len(changes)} pages updated\n")
    changes.sort(key=lambda r: (r[0], int(r[1])))
    for doc_id, page, old, new in changes:
        delta = (new - old) if old is not None else None
        if delta is None:
            print(f"  {doc_id} p.{page}: NEW = {new:.6f}")
        else:
            sign = "+" if delta >= 0 else ""
            print(f"  {doc_id} p.{page}: {old:.6f} -> {new:.6f} ({sign}{delta:.6f})")


if __name__ == "__main__":
    main()
