"""Fuzz Cluster v2 — filter multi-page mismatches before clustering."""
from __future__ import annotations
import json, sys
from pathlib import Path
from collections import defaultdict
from statistics import mean

sys.stdout.reconfigure(encoding='utf-8')


def flatten_attrs(meta_doc: dict) -> dict:
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


def main(batch_name: str, anchor_threshold: float = 50.0):
    batch_dir = Path(__file__).parent / "fuzz_runs" / batch_name
    divergences = json.loads((batch_dir / "divergences.json").read_text(encoding="utf-8"))

    clean = [d for d in divergences
             if d.get("anchor_diff") is not None
             and abs(d.get("anchor_diff", 0)) < anchor_threshold]
    print(f"Filtered: {len(clean)} / {len(divergences)} docs (|anchor_diff| < {anchor_threshold}pt)")

    attr_diffs = defaultdict(list)
    for d in clean:
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

    rows = []
    for (k, v), diffs in attr_diffs.items():
        if len(diffs) < 5:
            continue
        rows.append({
            "attr": k, "value": v, "n_docs": len(diffs),
            "mean_diff": mean(diffs), "max_diff": max(diffs),
        })
    rows.sort(key=lambda r: -r["mean_diff"])

    print(f"\n=== Top 30 attribute-value correlations (n_docs >= 5) ===")
    print(f"{'attr':<35} {'value':<25} {'n':>5} {'mean':>8} {'max':>8}")
    for r in rows[:30]:
        print(f"{r['attr']:<35} {r['value']:<25} {r['n_docs']:>5} {r['mean_diff']:>8.2f} {r['max_diff']:>8.2f}")

    print(f"\n=== Bottom 10 (well-matched) ===")
    for r in sorted(rows, key=lambda r: r["mean_diff"])[:10]:
        print(f"{r['attr']:<35} {r['value']:<25} {r['n_docs']:>5} {r['mean_diff']:>8.2f}")

    # Look for attr-value PAIRS (top pairs)
    pair_diffs = defaultdict(list)
    for d in clean:
        if d.get("error"): continue
        meta = d.get("meta")
        if not meta: continue
        flat = flatten_attrs(meta)
        flat_items = []
        for k, vals in flat.items():
            for v in vals:
                flat_items.append((k, repr(v)))
        # pairs
        diff = d.get("max_diff", 0)
        for i in range(len(flat_items)):
            for j in range(i+1, len(flat_items)):
                pair = (flat_items[i], flat_items[j])
                pair_diffs[pair].append(diff)

    pair_rows = []
    for pair, diffs in pair_diffs.items():
        if len(diffs) < 4: continue
        pair_rows.append({
            "p1": pair[0], "p2": pair[1], "n": len(diffs),
            "mean_diff": mean(diffs), "max_diff": max(diffs),
        })
    pair_rows.sort(key=lambda r: -r["mean_diff"])
    print(f"\n=== Top 15 attribute PAIRS (n >= 4) ===")
    for r in pair_rows[:15]:
        a1 = f"{r['p1'][0]}={r['p1'][1]}"[:30]
        a2 = f"{r['p2'][0]}={r['p2'][1]}"[:30]
        print(f"  {a1} × {a2}  n={r['n']} mean={r['mean_diff']:.2f}")

    (batch_dir / "cluster_clean.json").write_text(
        json.dumps({"singles": rows[:50], "pairs": pair_rows[:30]}, indent=2, default=str),
        encoding="utf-8"
    )


if __name__ == "__main__":
    batch = sys.argv[1] if len(sys.argv) > 1 else "alpha01"
    main(batch_name=batch)
