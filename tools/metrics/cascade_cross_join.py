"""Cross-join Word COM per-paragraph Y with Oxi --dump-layout per-paragraph Y.

Phase 1 Session 1 of cascade_unification_plan.md (commit 9e3638b),
follow-up to measure_cascade_y_diff.py.

Mismatch handled: Word COM enumerates 241 paragraphs (incl. table cells),
Oxi --dump-layout aggregates to ~56 distinct para_idx (body-level only).
We cannot index-match — we match by text content + page within tolerance.

Method:
  1. Word side: per-paragraph (page, y_pt, x_pt, text)
  2. Oxi side: aggregate text glyph elements by (page, para_idx)
       → top Y, top X, joined text (first 30 chars after combining frags)
  3. Match: for each Oxi paragraph with non-empty text, find Word paragraph
     with closest normalized text on the same page (or a nearby page if
     cascade has shifted it across the boundary)
  4. Compute delta_y = oxi_y - word_y, delta_page = oxi_page - word_page
  5. Sort by abs(delta_y), output top 30

Output: pipeline_data/cascade_y_diff/<doc_id>.json (per-doc detail)
        pipeline_data/cascade_y_diff/_summary.json (top sources across docs)

Run from repo root:
    python tools/metrics/cascade_cross_join.py 2ea81a
    python tools/metrics/cascade_cross_join.py            # all docs with both inputs
"""
from __future__ import annotations

import json
import os
import re
import sys
from collections import defaultdict

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
WORD_DIR = os.path.join(REPO_ROOT, "pipeline_data", "cascade_word_y")
OXI_DIR = os.path.join(REPO_ROOT, "pipeline_data", "cascade_oxi_y")
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "cascade_y_diff")


def normalize_text(s: str) -> str:
    """Strip whitespace + control chars; preserve printable content."""
    if not s:
        return ""
    s = s.replace("　", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def aggregate_oxi(oxi_data: dict, *, line_tol_pt: float = 0.5) -> list[dict]:
    """Group Oxi text elements into rendered "lines" by (page, rounded_Y).

    Includes elements with AND without para_idx — table cell text often
    has para_idx=None but is what we need for matching to Word's per-cell
    paragraph data.

    Lines are clustered by Y within ±line_tol_pt to absorb sub-pixel
    inconsistency in y-coordinate per glyph (rare but observed).
    """
    rows = []
    for pi, page in enumerate(oxi_data["pages"]):
        # Collect text elements, sort by (y, x)
        text_els = [e for e in page["elements"] if e["type"] == "text"]
        text_els.sort(key=lambda e: (e["y"], e["x"]))

        # Cluster by Y proximity — but split clusters when X resets backwards
        # (= a new line, even if same Y due to RTL or column wrap)
        clusters: list[list[dict]] = []
        cur: list[dict] = []
        cur_y: float | None = None
        cur_last_x: float = -1e9
        for e in text_els:
            if cur_y is None or abs(e["y"] - cur_y) > line_tol_pt or e["x"] < cur_last_x - 5.0:
                if cur:
                    clusters.append(cur)
                cur = [e]
                cur_y = e["y"]
                cur_last_x = e["x"]
            else:
                cur.append(e)
                cur_last_x = e["x"]
        if cur:
            clusters.append(cur)

        for cl in clusters:
            cl_sorted = sorted(cl, key=lambda e: e["x"])
            ys = [e["y"] for e in cl]
            xs = [e["x"] for e in cl]
            joined = "".join(e["text"] for e in cl_sorted)[:60]
            # Most common para_idx in the cluster (None counted separately)
            pidx_set = set(e["para_idx"] for e in cl)
            pidx = next(iter(pidx_set - {None}), None)  # any non-None
            rows.append({
                "page": pi + 1,
                "para_idx": pidx,
                "y_pt": min(ys),
                "x_pt": min(xs),
                "n_frags": len(cl),
                "text": joined,
                "text_norm": normalize_text(joined),
            })
    return rows


def normalize_word_paragraphs(word_data: dict) -> list[dict]:
    rows = []
    for r in word_data["paragraphs"]:
        rows.append({
            "i": r["i"],
            "page": r["page"],
            "y_pt": r["y_pt"],
            "x_pt": r["x_pt"],
            "text": r["text"],
            "text_norm": normalize_text(r["text"]),
            "in_table": r.get("in_table", False),
            "font": r.get("font", ""),
            "size_pt": r.get("size_pt"),
            "style": r.get("style", ""),
        })
    return rows


def match_paragraphs(oxi_rows: list[dict], word_rows: list[dict]) -> list[dict]:
    """For each Oxi row with non-trivial text, find best-matching Word row.

    Matching key: identical text_norm (first 20 chars) + page within ±1.
    If no exact match, fall back to longest-prefix match within ±2 pages.
    """
    matches = []
    used_word = set()  # word i indices already matched

    # Build text → list[word_row] index for fast prefix matching
    word_by_prefix: dict[str, list[dict]] = defaultdict(list)
    for w in word_rows:
        if not w["text_norm"]:
            continue
        prefix = w["text_norm"][:20]
        word_by_prefix[prefix].append(w)

    for o in oxi_rows:
        if not o["text_norm"] or len(o["text_norm"]) < 2:
            continue  # skip empty / single-char paras
        prefix = o["text_norm"][:20]
        candidates = word_by_prefix.get(prefix, [])

        # Filter: not already matched, page within ±2
        free = [w for w in candidates if w["i"] not in used_word]
        if not free:
            # Fallback: try shorter prefix (10 chars)
            short = o["text_norm"][:10]
            for w in word_rows:
                if w["i"] in used_word:
                    continue
                if w["text_norm"][:10] == short and short:
                    free.append(w)

        if not free:
            matches.append({
                "oxi_page": o["page"],
                "oxi_para_idx": o["para_idx"],
                "oxi_y": o["y_pt"],
                "oxi_text": o["text"],
                "word_i": None,
                "word_page": None,
                "word_y": None,
                "word_text": None,
                "delta_y": None,
                "delta_page": None,
                "in_table": None,
                "font": None,
                "size_pt": None,
                "style": None,
            })
            continue

        # Pick the closest by page
        free.sort(key=lambda w: (abs(w["page"] - o["page"]),
                                  abs((w["y_pt"] or 0) - o["y_pt"])))
        w = free[0]
        used_word.add(w["i"])
        matches.append({
            "oxi_page": o["page"],
            "oxi_para_idx": o["para_idx"],
            "oxi_y": o["y_pt"],
            "oxi_text": o["text"],
            "word_i": w["i"],
            "word_page": w["page"],
            "word_y": w["y_pt"],
            "word_text": w["text"],
            "delta_y": (o["y_pt"] - w["y_pt"]) if w["y_pt"] is not None else None,
            "delta_page": (o["page"] - w["page"]) if w["page"] is not None else None,
            "in_table": w["in_table"],
            "font": w["font"],
            "size_pt": w["size_pt"],
            "style": w["style"],
        })

    return matches


def summarize(matches: list[dict]) -> dict:
    n_total = len(matches)
    n_unmatched = sum(1 for m in matches if m["word_i"] is None)
    matched = [m for m in matches if m["word_i"] is not None]
    n_same_page = sum(1 for m in matched if m["delta_page"] == 0)
    n_page_shift = sum(1 for m in matched if m["delta_page"] != 0)
    deltas = [m["delta_y"] for m in matched if m["delta_y"] is not None and m["delta_page"] == 0]
    deltas_abs_sorted = sorted(matched, key=lambda m: abs(m["delta_y"]) if m["delta_y"] is not None else -1, reverse=True)

    # Bucket by in_table
    table_matches = [m for m in matched if m["in_table"]]
    body_matches = [m for m in matched if not m["in_table"]]

    return {
        "n_oxi_paragraphs": n_total,
        "n_unmatched": n_unmatched,
        "n_matched": len(matched),
        "n_same_page": n_same_page,
        "n_page_shift": n_page_shift,
        "table_match_count": len(table_matches),
        "body_match_count": len(body_matches),
        "delta_y_mean_same_page": (sum(deltas) / len(deltas)) if deltas else None,
        "delta_y_max_pos": max(deltas) if deltas else None,
        "delta_y_max_neg": min(deltas) if deltas else None,
        "top_30_by_abs_delta": [
            {
                "oxi_page": m["oxi_page"],
                "oxi_para_idx": m["oxi_para_idx"],
                "word_i": m["word_i"],
                "word_page": m["word_page"],
                "delta_page": m["delta_page"],
                "oxi_y": round(m["oxi_y"], 2) if m["oxi_y"] is not None else None,
                "word_y": round(m["word_y"], 2) if m["word_y"] is not None else None,
                "delta_y": round(m["delta_y"], 2) if m["delta_y"] is not None else None,
                "in_table": m["in_table"],
                "font": m["font"],
                "size_pt": m["size_pt"],
                "style": m["style"],
                "oxi_text": m["oxi_text"][:40],
                "word_text": m["word_text"][:40] if m["word_text"] else None,
            }
            for m in deltas_abs_sorted[:30]
        ],
    }


def process_doc(doc_id: str) -> dict | None:
    word_path = os.path.join(WORD_DIR, f"{doc_id}.json")
    oxi_path = os.path.join(OXI_DIR, f"{doc_id}.json")
    if not os.path.exists(word_path):
        print(f"  SKIP {doc_id}: no Word data ({word_path})")
        return None
    if not os.path.exists(oxi_path):
        print(f"  SKIP {doc_id}: no Oxi data ({oxi_path})")
        return None

    with open(word_path, encoding="utf-8") as f:
        word_data = json.load(f)
    with open(oxi_path, encoding="utf-8") as f:
        oxi_data = json.load(f)

    word_rows = normalize_word_paragraphs(word_data)
    oxi_rows = aggregate_oxi(oxi_data)
    matches = match_paragraphs(oxi_rows, word_rows)

    summary = summarize(matches)
    summary.update({
        "doc_id": doc_id,
        "filename": word_data.get("filename"),
        "floor_page": word_data.get("floor_page"),
        "floor_ssim": word_data.get("floor_ssim"),
        "issue": word_data.get("issue"),
        "n_pages": word_data.get("n_pages"),
        "n_word_paragraphs": word_data.get("n_paras"),
        "n_oxi_aggregated": len(oxi_rows),
    })

    out_path = os.path.join(OUT_DIR, f"{doc_id}.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"summary": summary, "matches": matches}, f, ensure_ascii=False, indent=2)
    return summary


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)

    if len(sys.argv) > 1:
        # Match prefix to known doc IDs
        target = sys.argv[1]
        docs_to_process = []
        for fname in os.listdir(WORD_DIR):
            if fname.endswith(".json") and not fname.startswith("_"):
                doc_id = fname[:-5]
                if doc_id.startswith(target):
                    docs_to_process.append(doc_id)
        if not docs_to_process:
            print(f"no Word data files matching prefix '{target}' in {WORD_DIR}")
            return 2
    else:
        docs_to_process = sorted(
            f[:-5] for f in os.listdir(WORD_DIR)
            if f.endswith(".json") and not f.startswith("_")
        )

    all_summaries = []
    for doc_id in docs_to_process:
        print(f"=== {doc_id} ===")
        summary = process_doc(doc_id)
        if summary is None:
            continue
        print(f"  oxi paras (with text): {summary['n_matched'] + summary['n_unmatched']}, "
              f"matched: {summary['n_matched']}, "
              f"unmatched: {summary['n_unmatched']}, "
              f"page-shift: {summary['n_page_shift']}")
        if summary["delta_y_mean_same_page"] is not None:
            print(f"  same-page delta_y: mean={summary['delta_y_mean_same_page']:+.2f}pt, "
                  f"max+={summary['delta_y_max_pos']:+.2f}, max-={summary['delta_y_max_neg']:+.2f}")
        if summary["top_30_by_abs_delta"]:
            print(f"  top 5 by abs(delta_y):")
            for r in summary["top_30_by_abs_delta"][:5]:
                d = r["delta_y"]
                dp = r["delta_page"]
                pg = r["oxi_page"]
                print(f"    p.{pg} dy={d:+.2f}pt dp={dp:+d} in_table={r['in_table']} "
                      f"{r['font']} {r['size_pt']}pt | "
                      f"{r['oxi_text']!r}")
        all_summaries.append(summary)
        print()

    summary_path = os.path.join(OUT_DIR, "_summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump({"docs": all_summaries}, f, ensure_ascii=False, indent=2)
    print(f"summary -> {summary_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
