"""Phase 2 element IoU diff: per-paragraph y-range IoU between Word and Oxi.

Phase 2 gate of the redesigned merge methodology (CLAUDE.md §Merge gate):
  - Per element (paragraph for now), compute Word vs Oxi bbox IoU
  - Aggregate: mean IoU per doc + cross-doc mean
  - Phase 2 → Phase 3 transition: mean IoU ≥ 0.99 sustained 5+ commits

Initial implementation: 1D IoU on paragraph y-range (start_y, end_y).
Y-range is more informative than full bbox for body paragraphs (x is
usually fixed by margins). Future extensions: full-bbox for table cells,
images, floating shapes.

Reuses existing inputs:
  - pipeline_data/pagination_word/<doc>.json (Word per-paragraph i, page, y, text)
  - pipeline_data/pagination_oxi/<doc>.json (Oxi per-page text records with y)

Per-paragraph height = next-paragraph-y - this-y (within same page) OR
default-line-height if last on page.

Output:
  pipeline_data/element_iou_diff/<doc>.json (per-paragraph IoU detail)
  pipeline_data/element_iou_diff/_summary.json (cross-doc mean IoU)

Run from repo root:
  python tools/metrics/element_iou_diff.py            # all docs with both inputs
  python tools/metrics/element_iou_diff.py 1636       # prefix filter
"""
from __future__ import annotations

import json
import os
import re
import sys

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
WORD_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_word")
OXI_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_oxi")
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "element_iou_diff")

MIN_MATCH_LEN = 2
PAGE_SEARCH_RADIUS = 2  # narrower than pagination_diff: IoU only meaningful on same page
DEFAULT_LINE_H = 14.0   # fallback when paragraph has no following paragraph


def normalize_text(s: str) -> str:
    if not s:
        return ""
    s = s.replace("　", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def yrange_iou(s1: float, e1: float, s2: float, e2: float) -> float:
    """1D Intersection-over-Union on two y-ranges."""
    inter = max(0.0, min(e1, e2) - max(s1, s2))
    union = max(e1, e2) - min(s1, s2)
    if union <= 0.0:
        return 0.0
    return inter / union


def load_word(doc_id: str) -> dict | None:
    p = os.path.join(WORD_DIR, f"{doc_id}.json")
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def load_oxi(doc_id: str) -> dict | None:
    p = os.path.join(OXI_DIR, f"{doc_id}.json")
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def derive_word_heights(paragraphs: list[dict]) -> list[dict]:
    """For each Word paragraph, derive height from next-paragraph y diff
    (within same page). For last on page, use DEFAULT_LINE_H."""
    out = []
    for i, p in enumerate(paragraphs):
        if p.get("y") is None or p.get("page") is None:
            continue
        h = None
        for j in range(i + 1, len(paragraphs)):
            np = paragraphs[j]
            if np.get("page") != p["page"]:
                break
            if np.get("y") is None:
                continue
            if np["y"] > p["y"]:
                h = np["y"] - p["y"]
                break
        if h is None or h <= 0:
            h = DEFAULT_LINE_H
        out.append({**p, "h": h, "y_end": p["y"] + h})
    return out


def derive_oxi_heights(pages: dict[str, list[dict]]) -> list[dict]:
    """For each Oxi paragraph, derive height from next-paragraph y diff
    (within same page). Returns a flat list with page info."""
    out = []
    for page_str in sorted(pages.keys(), key=int):
        page = int(page_str)
        recs = pages[page_str]
        # Sort by y
        sorted_recs = sorted(recs, key=lambda r: r.get("y", 0))
        for i, r in enumerate(sorted_recs):
            if r.get("y") is None:
                continue
            h = None
            for j in range(i + 1, len(sorted_recs)):
                nr = sorted_recs[j]
                if nr.get("y") is None:
                    continue
                if nr["y"] > r["y"]:
                    h = nr["y"] - r["y"]
                    break
            if h is None or h <= 0:
                h = DEFAULT_LINE_H
            out.append({**r, "page": page, "h": h, "y_end": r["y"] + h})
    return out


def diff_doc(doc_id: str, word: dict, oxi: dict) -> dict:
    word_paras = derive_word_heights(word.get("paragraphs", []))
    oxi_paras = derive_oxi_heights(oxi.get("pages", {}))

    # Build (page → list of (text_norm, record)) for Oxi for matching
    oxi_by_page: dict[int, list[tuple[str, dict]]] = {}
    for r in oxi_paras:
        t = normalize_text(r.get("text", ""))
        if len(t) < MIN_MATCH_LEN:
            continue
        oxi_by_page.setdefault(r["page"], []).append((t, r))

    used: set[tuple[int, int]] = set()
    matches: list[dict] = []

    for wp in word_paras:
        wt = normalize_text(wp.get("text", ""))
        if len(wt) < MIN_MATCH_LEN:
            continue
        wpage = wp["page"]
        # Search same page first, then expand
        best = None
        best_dist = None
        for radius in range(0, PAGE_SEARCH_RADIUS + 1):
            for sign in ((0,) if radius == 0 else (-1, +1)):
                cand_page = wpage + sign * radius
                if cand_page < 1:
                    continue
                cand = oxi_by_page.get(cand_page, [])
                for idx, (ot, _orec) in enumerate(cand):
                    if (cand_page, idx) in used:
                        continue
                    n = min(len(wt), len(ot))
                    if n < MIN_MATCH_LEN:
                        continue
                    if wt[:n] == ot[:n]:
                        dist = (radius, -n)
                        if best is None or dist < best_dist:
                            best = (cand_page, idx)
                            best_dist = dist
            if best is not None and radius == 0:
                break

        if best is None:
            matches.append({
                "word_i": wp["i"],
                "word_page": wpage,
                "word_y": wp["y"],
                "word_h": wp["h"],
                "oxi_y": None,
                "oxi_h": None,
                "iou": None,
                "matched": False,
            })
            continue

        used.add(best)
        opage, oidx = best
        orec = oxi_by_page[opage][oidx][1]
        # Only compute IoU when on same page (cross-page IoU is degenerate)
        if opage == wpage:
            iou = yrange_iou(wp["y"], wp["y_end"], orec["y"], orec["y_end"])
        else:
            iou = 0.0  # different page = zero overlap
        matches.append({
            "word_i": wp["i"],
            "word_page": wpage,
            "word_y": round(wp["y"], 2),
            "word_h": round(wp["h"], 2),
            "oxi_page": opage,
            "oxi_y": round(orec["y"], 2),
            "oxi_h": round(orec["h"], 2),
            "iou": round(iou, 4),
            "matched": True,
        })

    matched = [m for m in matches if m["matched"]]
    n_matched = len(matched)
    if n_matched > 0:
        mean_iou = sum(m["iou"] for m in matched) / n_matched
        n_iou_high = sum(1 for m in matched if m["iou"] >= 0.99)
        n_iou_zero = sum(1 for m in matched if m["iou"] == 0.0)
    else:
        mean_iou = 0.0
        n_iou_high = 0
        n_iou_zero = 0

    return {
        "doc_id": doc_id,
        "n_word_paras": len(word_paras),
        "n_matched": n_matched,
        "n_unmatched": len(matches) - n_matched,
        "mean_iou": round(mean_iou, 4),
        "n_iou_high": n_iou_high,  # paragraphs with IoU ≥ 0.99 (Phase 2 criterion per-element)
        "n_iou_zero": n_iou_zero,  # cross-page or no-overlap
        "frac_iou_high": round(n_iou_high / n_matched, 4) if n_matched else 0.0,
        "matches": matches,
    }


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)

    prefix = sys.argv[1] if len(sys.argv) > 1 else None
    word_files = sorted(f for f in os.listdir(WORD_DIR)
                        if f.endswith(".json") and not f.startswith("_"))
    doc_ids = [os.path.splitext(f)[0] for f in word_files]
    if prefix:
        doc_ids = [d for d in doc_ids if d.startswith(prefix)]
    if not doc_ids:
        print(f"no docs matched (prefix={prefix})", file=sys.stderr)
        return 2

    summary = []
    skipped = []
    for doc_id in doc_ids:
        w = load_word(doc_id)
        o = load_oxi(doc_id)
        if w is None or o is None:
            skipped.append({"doc_id": doc_id, "has_word": w is not None, "has_oxi": o is not None})
            continue
        result = diff_doc(doc_id, w, o)
        out_path = os.path.join(OUT_DIR, f"{doc_id}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"  {doc_id}: mean_iou={result['mean_iou']:.4f} n_high={result['n_iou_high']}/{result['n_matched']} ({result['frac_iou_high']:.1%})")
        summary.append({
            "doc_id": doc_id,
            "mean_iou": result["mean_iou"],
            "n_matched": result["n_matched"],
            "n_unmatched": result["n_unmatched"],
            "n_iou_high": result["n_iou_high"],
            "frac_iou_high": result["frac_iou_high"],
        })

    n = len(summary)
    overall_mean = sum(s["mean_iou"] for s in summary) / n if n else 0.0
    n_pass = sum(1 for s in summary if s["mean_iou"] >= 0.99)
    summary_obj = {
        "n_total": n,
        "n_pass": n_pass,
        "n_fail": n - n_pass,
        "pass_rate": round(n_pass / n, 4) if n else 0.0,
        "mean_iou": round(overall_mean, 4),
        "n_skipped": len(skipped),
        "docs": summary,
        "skipped": skipped,
    }
    out_path = os.path.join(OUT_DIR, "_summary.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(summary_obj, f, ensure_ascii=False, indent=2)
    print(f"\nsummary -> {out_path}")
    if n:
        print(f"  PHASE 2 GATE: mean_iou={overall_mean:.4f} pass(>=0.99)={n_pass}/{n} ({n_pass/n:.1%})")
    else:
        print("  no docs")
    if skipped:
        print(f"  skipped {len(skipped)} docs (missing input)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
