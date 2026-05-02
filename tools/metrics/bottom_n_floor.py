"""
Compute bottom-N min(SSIM) per doc and the floor sum.

Usage:
  python tools/metrics/bottom_n_floor.py [path/to/ssim.json] [N]

Default: pipeline_data/ssim_baseline.json, N=5.
Prints sorted ascending list and the bottom-N sum.
"""
import json
import os
import sys


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        "pipeline_data", "ssim_baseline.json"
    )
    n = int(sys.argv[2]) if len(sys.argv) > 2 else 5
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    rows = []
    for doc_id, pages in data.items():
        if not pages:
            continue
        m = min(float(v) for v in pages.values())
        rows.append((m, doc_id))
    rows.sort()
    print(f"# {path}: {len(rows)} docs, N={n}\n")
    for m, doc in rows[: max(n, 10)]:
        print(f"{m:.6f}  {doc}")
    floor_sum = sum(m for m, _ in rows[:n])
    print(f"\nbottom-{n} floor sum: {floor_sum:.6f}")
    print(f"mean min(SSIM): {sum(m for m,_ in rows)/len(rows):.6f}")


if __name__ == "__main__":
    main()
