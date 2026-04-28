"""Phase 2 Session 1 of cascade_unification_plan.md — find compression source CLASSES.

Cross-join produces per-paragraph delta_y, but absolute deltas mix cumulative
upstream errors with the local paragraph's height error. To isolate per-paragraph
compression, this analyzer:

  1. Computes "step delta" (incremental) between consecutive matched
     paragraphs on the same page in Word's reading order:
        word_step = word_y(i+1) - word_y(i)
        oxi_step  = oxi_y(i+1)  - oxi_y(i)
        step_delta = oxi_step - word_step
     A negative step_delta means Oxi's row was SHORTER than Word's
     (compression). Cumulative effects from earlier paragraphs cancel out.

  2. Buckets step deltas by characteristics of the FIRST paragraph in each
     pair (in_table × style × font_family × size_range), reporting per
     bucket: count, mean step_delta, stddev, top |step_delta| examples.

  3. Cross-doc: aggregates the same buckets across all 5 floor docs to
     reveal compression source CLASSES that are universal vs doc-specific.

Output: pipeline_data/cascade_y_diff/_pattern_analysis.json (machine-readable)
         + console summary of top buckets by abs(mean step_delta) × count.

Run from repo root:
    python tools/metrics/cascade_pattern_analyzer.py
"""
from __future__ import annotations

import json
import math
import os
import sys
from collections import defaultdict

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DIFF_DIR = os.path.join(REPO_ROOT, "pipeline_data", "cascade_y_diff")
OUT_PATH = os.path.join(DIFF_DIR, "_pattern_analysis.json")


def size_bucket(size: float | None) -> str:
    if size is None:
        return "?"
    if size < 8:
        return "<8"
    if size < 10:
        return "8-10"
    if size < 11:
        return "10-11"
    if size < 12:
        return "11-12"
    if size < 14:
        return "12-14"
    if size < 16:
        return "14-16"
    return ">=16"


def font_family(name: str | None) -> str:
    if not name:
        return "?"
    n = name.strip()
    # Common Japanese font normalization
    for prefix, key in [
        ("ＭＳ", "MS"),
        ("MS ", "MS"),
        ("Yu Gothic", "Yu Gothic"),
        ("游ゴシック", "Yu Gothic"),
        ("Yu Mincho", "Yu Mincho"),
        ("游明朝", "Yu Mincho"),
        ("Meiryo", "Meiryo"),
        ("メイリオ", "Meiryo"),
        ("BIZ UDP", "BIZ UDP"),
        ("Noto", "Noto"),
        ("Times New Roman", "Times New Roman"),
        ("Calibri", "Calibri"),
        ("Cambria", "Cambria"),
        ("Century", "Century"),
        ("Arial", "Arial"),
        ("ヒラギノ", "Hiragino"),
    ]:
        if n.startswith(prefix):
            return key
    return n[:20]


def compute_steps(matches: list[dict]) -> list[dict]:
    """For each consecutive same-page matched pair (in Word order), produce a step record.

    Word order is by word_i (paragraph index in document order).
    """
    matched = [m for m in matches if m["word_i"] is not None]
    matched.sort(key=lambda m: (m["word_page"] or 0, m["word_i"]))

    steps = []
    for i in range(len(matched) - 1):
        a = matched[i]
        b = matched[i + 1]
        if a["word_page"] != b["word_page"]:
            continue  # cross-page step is dominated by margin/page-break, not single paragraph
        if a["delta_page"] != 0 or b["delta_page"] != 0:
            continue  # skip cascade-shifted pairs
        if a["word_y"] is None or b["word_y"] is None:
            continue
        if a["oxi_y"] is None or b["oxi_y"] is None:
            continue
        word_step = b["word_y"] - a["word_y"]
        oxi_step = b["oxi_y"] - a["oxi_y"]
        step_delta = oxi_step - word_step
        if word_step < 0:
            continue  # paragraphs out of Y-order = noise (shouldn't happen)
        steps.append({
            "word_i_a": a["word_i"],
            "word_i_b": b["word_i"],
            "page": a["word_page"],
            "word_step_pt": word_step,
            "oxi_step_pt": oxi_step,
            "step_delta_pt": step_delta,
            "in_table_a": a["in_table"],
            "in_table_b": b["in_table"],
            "style_a": a["style"],
            "font_a": a["font"],
            "size_a": a["size_pt"],
            "text_a": a["oxi_text"][:40],
            "text_b": b["oxi_text"][:40],
        })
    return steps


def bucket_key(step: dict) -> tuple:
    """Bucket the step by characteristics of paragraph A (the source of the row height)."""
    in_t = "cell" if step["in_table_a"] else "body"
    style = step["style_a"] or "?"
    font = font_family(step["font_a"])
    size = size_bucket(step["size_a"])
    return (in_t, style, font, size)


def stats(values: list[float]) -> dict:
    if not values:
        return {"n": 0, "mean": None, "std": None, "min": None, "max": None}
    n = len(values)
    mean = sum(values) / n
    var = sum((v - mean) ** 2 for v in values) / n
    std = math.sqrt(var)
    return {
        "n": n,
        "mean": round(mean, 3),
        "std": round(std, 3),
        "min": round(min(values), 3),
        "max": round(max(values), 3),
    }


def analyze(docs: list[dict]) -> dict:
    # Per-doc steps
    per_doc = {}
    cross_buckets = defaultdict(lambda: {"docs": defaultdict(list), "all": []})

    for doc in docs:
        doc_id = doc["doc_id"]
        steps = doc["steps"]

        # Per-bucket aggregation within doc
        buckets = defaultdict(list)
        for s in steps:
            buckets[bucket_key(s)].append(s)

        per_doc[doc_id] = {
            "floor_page": doc.get("floor_page"),
            "n_steps": len(steps),
            "overall": stats([s["step_delta_pt"] for s in steps]),
            "by_in_table": {
                "cell": stats([s["step_delta_pt"] for s in steps if s["in_table_a"]]),
                "body": stats([s["step_delta_pt"] for s in steps if not s["in_table_a"]]),
            },
            "top_buckets": sorted(
                [
                    {
                        "key": list(k),
                        "stats": stats([s["step_delta_pt"] for s in v]),
                    }
                    for k, v in buckets.items()
                    if len(v) >= 3
                ],
                key=lambda b: (abs(b["stats"]["mean"] or 0) * b["stats"]["n"]),
                reverse=True,
            )[:10],
        }

        # Cross-doc accumulation
        for k, v in buckets.items():
            for s in v:
                cross_buckets[k]["docs"][doc_id].append(s["step_delta_pt"])
                cross_buckets[k]["all"].append(s["step_delta_pt"])

    # Cross-doc ranking
    cross_ranked = []
    for k, v in cross_buckets.items():
        all_deltas = v["all"]
        if len(all_deltas) < 5:
            continue
        st = stats(all_deltas)
        n_docs = sum(1 for d in v["docs"].values() if d)
        cross_ranked.append({
            "bucket": list(k),
            "n_docs_present": n_docs,
            "stats": st,
            "per_doc_means": {
                doc_id: round(sum(vs) / len(vs), 3) for doc_id, vs in v["docs"].items() if vs
            },
            "impact_score": round(abs(st["mean"] or 0) * st["n"], 1),
        })
    cross_ranked.sort(key=lambda b: b["impact_score"], reverse=True)

    return {"per_doc": per_doc, "cross_doc_ranked": cross_ranked}


def main() -> int:
    if not os.path.isdir(DIFF_DIR):
        print(f"missing {DIFF_DIR}; run cascade_cross_join.py first", file=sys.stderr)
        return 2

    docs = []
    for fname in sorted(os.listdir(DIFF_DIR)):
        if not fname.endswith(".json") or fname.startswith("_"):
            continue
        path = os.path.join(DIFF_DIR, fname)
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
        doc_id = fname[:-5]
        steps = compute_steps(d["matches"])
        docs.append({
            "doc_id": doc_id,
            "floor_page": d["summary"].get("floor_page"),
            "steps": steps,
        })
        print(f"  {doc_id}: {len(steps)} same-page step pairs")

    result = analyze(docs)

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print()
    print("=== Per-doc body vs cell summary ===")
    for doc_id, info in result["per_doc"].items():
        b = info["by_in_table"]["body"]
        c = info["by_in_table"]["cell"]
        print(f"  {doc_id}: overall n={info['overall']['n']:3} mean={info['overall']['mean']:+6.2f}  "
              f"body[n={b['n']:3} mean={b['mean'] or 0:+6.2f} std={b['std'] or 0:5.2f}]  "
              f"cell[n={c['n']:3} mean={c['mean'] or 0:+6.2f} std={c['std'] or 0:5.2f}]")

    print()
    print("=== Top 15 cross-doc buckets by impact (|mean|×n) ===")
    print(f"  {'rank':<4} {'docs':<5} {'n':<5} {'mean':>8} {'std':>6} | {'in_t':<5} {'style':<20} {'font':<15} {'size':<6}")
    for i, b in enumerate(result["cross_doc_ranked"][:15], 1):
        s = b["stats"]
        k = b["bucket"]
        style_short = (k[1] or "?")[:18]
        font_short = (k[2] or "?")[:13]
        print(f"  {i:<4} {b['n_docs_present']:<5} {s['n']:<5} {s['mean']:>+8.2f} {s['std']:>6.2f} | "
              f"{k[0]:<5} {style_short:<20} {font_short:<15} {k[3]:<6}")

    print()
    print(f"Detail JSON -> {OUT_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
