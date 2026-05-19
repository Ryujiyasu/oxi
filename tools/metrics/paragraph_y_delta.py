"""Per-paragraph y-coordinate drift diagnostic.

For a given doc, text-prefix-match Word's per-paragraph y-data against
Oxi's pagination_oxi records. For each matched pair compute:

- page-relative y in Word vs Oxi
- "absolute" cumulative y = (page - 1) * page_height + y
  (page_height defaults to 842pt = A4; can be doc-specific)
- delta = Oxi - Word
- d_delta = delta_i - delta_{i-1} (per-paragraph drift CONTRIBUTION)

Output: per-paragraph table sorted by |d_delta| descending (drift origins)
and a chronological table (drift evolution across the doc).

Pair this with paragraph_y_delta_plot.py (TODO) to visualize cumulative
drift over paragraph index.

Usage:
  python tools/metrics/paragraph_y_delta.py <doc_id> [--top=20] [--limit=200]
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import unicodedata

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
WORD_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_word")
OXI_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_oxi")

# Same normalization as pagination_diff.py
def normalize_text(s: str) -> str:
    if not s:
        return ""
    out = []
    for ch in s:
        if ch.isspace() or ch == "　":
            continue
        out.append(ch)
    return unicodedata.normalize("NFKC", "".join(out))

MIN_MATCH_LEN = 6


def absolute_y(page: int, y: float, page_h: float = 842.0) -> float:
    """Cumulative y measured from page 1 top."""
    return (page - 1) * page_h + y


def match_pairs(word: dict, oxi: dict, page_h: float = 842.0) -> list[dict]:
    """Returns list of matched pairs sorted by Word paragraph index.

    Each pair: {word_i, word_page, word_y, oxi_page, oxi_y, text_norm,
                word_abs_y, oxi_abs_y, delta_y}.
    """
    # Build Oxi index: page -> list of (text_norm, y, record)
    oxi_by_page: dict[int, list[tuple[str, float, dict]]] = {}
    for page_str, recs in oxi["pages"].items():
        page = int(page_str)
        for r in recs:
            t = normalize_text(r.get("text", ""))
            if len(t) < MIN_MATCH_LEN:
                continue
            oxi_by_page.setdefault(page, []).append((t, r["y"], r))

    used: dict[tuple[int, int], int] = {}
    pairs = []
    for wp in word.get("paragraphs", []):
        wt = normalize_text(wp.get("text", ""))
        if len(wt) < MIN_MATCH_LEN:
            continue
        wpage = wp.get("page")
        wy = wp.get("y")
        if wpage is None or wy is None:
            continue

        # Search expected page first, then ±1, ±2, ±3
        best = None
        best_dist = None
        for radius in range(0, 4):
            for sign in ((0,) if radius == 0 else (-1, +1)):
                cand_page = wpage + sign * radius
                if cand_page < 1:
                    continue
                cand_recs = oxi_by_page.get(cand_page, [])
                for idx, (ot, oy, _r) in enumerate(cand_recs):
                    n = min(len(wt), len(ot))
                    if n < MIN_MATCH_LEN:
                        continue
                    if wt[:n] != ot[:n]:
                        continue
                    capacity = max(1, ot.count(wt[:n]))
                    if used.get((cand_page, idx), 0) >= capacity:
                        continue
                    dist = (radius, -n)
                    if best is None or dist < best_dist:
                        best = (cand_page, idx, oy)
                        best_dist = dist
            if best is not None and radius == 0:
                break

        if best is None:
            continue
        used[(best[0], best[1])] = used.get((best[0], best[1]), 0) + 1
        opage, oidx, oy = best
        pairs.append({
            "word_i": wp["i"],
            "word_page": wpage,
            "word_y": wy,
            "oxi_page": opage,
            "oxi_y": oy,
            "text": wt[:40],
            "word_abs_y": absolute_y(wpage, wy, page_h),
            "oxi_abs_y": absolute_y(opage, oy, page_h),
        })

    pairs.sort(key=lambda p: p["word_i"])
    # Compute delta and d_delta
    prev_delta = 0.0
    for p in pairs:
        p["delta_y"] = p["oxi_abs_y"] - p["word_abs_y"]
        p["d_delta"] = p["delta_y"] - prev_delta
        prev_delta = p["delta_y"]
    return pairs


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("doc_id", help="doc id prefix (e.g. 3a4f9fbe1a83)")
    ap.add_argument("--top", type=int, default=20,
                    help="top N drift contributions to print")
    ap.add_argument("--limit", type=int, default=200,
                    help="max chronological rows to print")
    ap.add_argument("--page-h", type=float, default=842.0,
                    help="page height pt (default A4=842)")
    ap.add_argument("--threshold", type=float, default=2.0,
                    help="d_delta threshold for 'jump' detection")
    ap.add_argument("--json", action="store_true",
                    help="also write detailed JSON output")
    args = ap.parse_args()

    word_path = os.path.join(WORD_DIR, f"{args.doc_id}.json")
    oxi_path = os.path.join(OXI_DIR, f"{args.doc_id}.json")
    if not os.path.exists(word_path):
        # try prefix match
        for fn in os.listdir(WORD_DIR):
            if fn.startswith(args.doc_id):
                word_path = os.path.join(WORD_DIR, fn)
                break
    if not os.path.exists(oxi_path):
        for fn in os.listdir(OXI_DIR):
            if fn.startswith(args.doc_id):
                oxi_path = os.path.join(OXI_DIR, fn)
                break

    with open(word_path, encoding="utf-8") as f:
        word = json.load(f)
    with open(oxi_path, encoding="utf-8") as f:
        oxi = json.load(f)

    pairs = match_pairs(word, oxi, args.page_h)

    # Stats
    n = len(pairs)
    final_delta = pairs[-1]["delta_y"] if pairs else 0.0
    print(f"=== {args.doc_id}: {n} matched paragraphs ===")
    print(f"final cumulative y-delta (Oxi - Word): {final_delta:+.2f}pt")
    print(f"page_h assumed = {args.page_h}pt\n")

    # Top jumps (largest |d_delta|)
    print(f"=== TOP {args.top} DRIFT JUMPS (|d_delta| desc) ===")
    print("idx_in_match  word_i  word_p  oxi_p  d_delta   text")
    by_jump = sorted(enumerate(pairs), key=lambda t: -abs(t[1]["d_delta"]))
    for idx, p in by_jump[:args.top]:
        print(f"  {idx:5d}      {p['word_i']:5d}  p.{p['word_page']:3d}  p.{p['oxi_page']:3d}  {p['d_delta']:+7.2f}  {p['text']}")

    # Cumulative drift threshold crossings
    print()
    print(f"=== JUMPS > |{args.threshold}|pt CHRONOLOGICAL ===")
    print("idx_in_match  word_i  page  cumulative_delta_y  d_delta   text")
    for idx, p in enumerate(pairs):
        if abs(p["d_delta"]) > args.threshold:
            print(f"  {idx:5d}      {p['word_i']:5d}  p.{p['word_page']:3d}->{p['oxi_page']:3d}  cum={p['delta_y']:+8.2f}  d={p['d_delta']:+7.2f}  {p['text']}")

    if args.json:
        out_path = os.path.join(REPO_ROOT, "tools", "metrics",
                                f"paragraph_y_delta_{args.doc_id}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump({"doc_id": args.doc_id, "n_matched": n,
                       "final_delta_y": final_delta, "pairs": pairs},
                      f, ensure_ascii=False, indent=2)
        print(f"\nJSON written to {out_path}")


if __name__ == "__main__":
    main()
