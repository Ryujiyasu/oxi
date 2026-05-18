"""Fuzz Quirk Discovery — attribute cluster analyzer.

For high-divergence docs, find common attribute patterns:
- Which attribute value is over-represented in high-diff vs low-diff?
- Score each attribute by correlation with divergence magnitude.

Output: ranked attribute -> divergence correlation table.
"""
from __future__ import annotations
import json
import sys
from pathlib import Path
from collections import defaultdict
from statistics import mean, stdev

sys.stdout.reconfigure(encoding='utf-8')


def flatten_attrs(meta_doc: dict) -> dict:
    """Flatten a doc's meta into a flat attr dict: attr -> set of values seen."""
    attrs = defaultdict(set)
    for row in meta_doc.get("rows", []):
        for k in ("tr_height", "h_rule", "cant_split"):
            v = row.get(k)
            if v is not None:
                attrs[f"row.{k}"].add(v)
        for cell in row.get("cells", []):
            for k in ("width", "grid_span", "v_align", "mar_top", "mar_bottom"):
                v = cell.get(k)
                if v is not None:
                    attrs[f"cell.{k}"].add(v)
            for p in cell.get("paragraphs", []):
                for k, v in p.items():
                    if v is not None:
                        attrs[f"para.{k}"].add(v)
    return {k: list(v) for k, v in attrs.items()}


def analyze_batch(batch_name: str) -> None:
    batch_dir = Path(__file__).parent / "fuzz_runs" / batch_name
    divergences = json.loads((batch_dir / "divergences.json").read_text(encoding="utf-8"))

    # Build (attr_key, attr_value) -> [diffs] map
    attr_diffs: dict[tuple[str, str], list[float]] = defaultdict(list)
    for d in divergences:
        if d.get("error"):
            continue
        diff = d.get("max_diff", 0)
        meta = d.get("meta")
        if not meta:
            continue
        flat = flatten_attrs(meta)
        for k, values in flat.items():
            for v in values:
                attr_diffs[(k, repr(v))].append(diff)

    # For each (attr, value), compute mean diff
    rows = []
    for (k, v), diffs in attr_diffs.items():
        if len(diffs) < 2:  # need at least 2 samples
            continue
        rows.append({
            "attr": k,
            "value": v,
            "n_docs": len(diffs),
            "mean_diff": mean(diffs),
            "max_diff": max(diffs),
        })

    rows.sort(key=lambda r: -r["mean_diff"])

    print(f"=== Top 25 attributes correlated with high divergence (batch: {batch_name}) ===")
    print(f"{'attr':<35} {'value':<25} {'n_docs':>6} {'mean':>8} {'max':>8}")
    for r in rows[:25]:
        print(f"{r['attr']:<35} {r['value']:<25} {r['n_docs']:>6} {r['mean_diff']:>8.2f} {r['max_diff']:>8.2f}")

    # Also: lowest mean_diff (i.e. attrs/values where Oxi matches Word)
    rows_clean = sorted(rows, key=lambda r: r["mean_diff"])
    print(f"\n=== Bottom 10 (matches Word well) ===")
    for r in rows_clean[:10]:
        print(f"{r['attr']:<35} {r['value']:<25} {r['n_docs']:>6} {r['mean_diff']:>8.2f}")

    out = batch_dir / "cluster.json"
    out.write_text(json.dumps(rows, indent=2, default=str), encoding="utf-8")
    print(f"\nFull cluster data: {out}")


if __name__ == "__main__":
    batch = sys.argv[1] if len(sys.argv) > 1 else "smoke"
    analyze_batch(batch_name=batch)
