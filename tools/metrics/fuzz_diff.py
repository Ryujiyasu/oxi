"""Fuzz Quirk Discovery — divergence analyzer.

Compares Word vs Oxi measurements across all fuzz docs. Output:
- Per-doc max divergence
- Top docs by divergence magnitude
- Attribute-cluster analysis: which attribute combinations produce divergence

The Word and Oxi y-conventions differ (~1.75pt offset for body paragraphs),
so we compare DELTAS within a doc (anchor_y - cell_y, etc) not absolute.
"""
from __future__ import annotations
import json
import sys
from pathlib import Path
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')


def doc_divergence(d: dict) -> dict:
    """For a single measured doc, compute Word vs Oxi divergence."""
    word = d.get("word", {})
    oxi = d.get("oxi", {})
    if "error" in word or "error" in oxi:
        return {"doc": d["doc"], "error": True}

    # head_y - cell_y (delta from head to each row's first cell)
    wh = word.get("head_y")
    oh = oxi.get("head_y")
    wa = word.get("anchor_y")
    oa = oxi.get("anchor_y")

    word_cells = []  # (row_idx, col_idx, y - head_y)
    for row in word.get("rows", []):
        for cell in row.get("cells", []):
            word_cells.append((row["row_idx"], cell["col_idx"], cell["y"] - wh if wh else None))
    oxi_cells = []
    for cell in oxi.get("rows", []):
        oxi_cells.append((cell["row_idx"], cell["col_idx"], cell["y"] - oh if oh else None))

    word_map = {(r, c): y for r, c, y in word_cells}
    oxi_map = {(r, c): y for r, c, y in oxi_cells}

    diffs = []
    keys = set(word_map.keys()) | set(oxi_map.keys())
    for key in sorted(keys):
        w_y = word_map.get(key)
        o_y = oxi_map.get(key)
        if w_y is None or o_y is None:
            diffs.append({"key": key, "missing": True, "word_y": w_y, "oxi_y": o_y})
            continue
        diff = w_y - o_y
        diffs.append({"key": key, "diff": diff, "word_dy": w_y, "oxi_dy": o_y})

    # anchor delta
    anchor_diff = None
    if wh and oh and wa and oa:
        anchor_diff = (wa - wh) - (oa - oh)

    max_diff = max((abs(d.get("diff", 0)) for d in diffs if "diff" in d), default=0)
    return {
        "doc": d["doc"],
        "max_diff": max_diff,
        "anchor_diff": anchor_diff,
        "diffs": diffs,
    }


def analyze_batch(batch_name: str) -> None:
    batch_dir = Path(__file__).parent / "fuzz_runs" / batch_name
    meas = json.loads((batch_dir / "measurements.json").read_text(encoding="utf-8"))
    meta = json.loads((batch_dir / "meta.json").read_text(encoding="utf-8"))
    meta_by_doc = {m["doc_name"]: m for m in meta}

    results = []
    for d in meas:
        r = doc_divergence(d)
        r["meta"] = meta_by_doc.get(d["doc"])
        results.append(r)

    # Sort by max_diff descending
    sorted_results = sorted(results, key=lambda x: -x.get("max_diff", 0))

    print(f"\n=== Top 10 divergences in batch '{batch_name}' ===")
    for r in sorted_results[:10]:
        if r.get("error"):
            print(f"  {r['doc']}: ERROR")
            continue
        print(f"  {r['doc']}: max_diff={r['max_diff']:.2f}pt, anchor_diff={r.get('anchor_diff', 'n/a')}")
        # show top 3 cell diffs
        cell_diffs = sorted(
            [d for d in r["diffs"] if "diff" in d],
            key=lambda x: -abs(x["diff"])
        )[:3]
        for cd in cell_diffs:
            print(f"    cell {cd['key']}: word_dy={cd['word_dy']:.2f} oxi_dy={cd['oxi_dy']:.2f} diff={cd['diff']:+.2f}")
        # show key attrs from meta
        if r.get("meta"):
            interesting = []
            for row_meta in r["meta"].get("rows", []):
                row_info = {
                    "row_idx": row_meta["row_idx"],
                    "trH": row_meta.get("tr_height"),
                    "hRule": row_meta.get("h_rule"),
                }
                for cell_meta in row_meta.get("cells", []):
                    for p in cell_meta.get("paragraphs", []):
                        keys_set = {k: v for k, v in p.items() if v is not None and v != 0}
                        if keys_set:
                            interesting.append((row_info, keys_set))
            if interesting:
                print(f"    attrs (first):", interesting[0][1] if interesting else "{}")

    # Distribution
    print(f"\n=== Divergence distribution ===")
    tiers = [(0, 1), (1, 2), (2, 5), (5, 10), (10, 100)]
    for lo, hi in tiers:
        count = sum(1 for r in results if lo <= r.get("max_diff", 0) < hi)
        print(f"  {lo}-{hi}pt: {count} docs")

    # Save full diff
    out_path = batch_dir / "divergences.json"
    out_path.write_text(json.dumps(sorted_results, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    print(f"\nFull divergences written to {out_path}")


if __name__ == "__main__":
    batch = sys.argv[1] if len(sys.argv) > 1 else "smoke"
    analyze_batch(batch_name=batch)
