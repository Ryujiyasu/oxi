"""Phase 2 element IoU diff: per-paragraph y-range IoU between Word and Oxi.

Phase 2 gate of the redesigned merge methodology (CLAUDE.md §Merge gate):
  - Per element (paragraph for now), compute Word vs Oxi bbox IoU
  - Aggregate: mean IoU per doc + cross-doc mean
  - Phase 2 → Phase 3 transition: mean IoU ≥ 0.99 sustained 5+ commits

R54 (2026-05-17): per-(in_table) median_dy bimodal repair. Word reports
table cell paragraph y at a different convention than body paragraphs
(typically +0.5pt offset, observable across 7+ baseline docs). Computing
a single doc-wide median leaves the minority cohort with systematic
±0.5pt residual that is convention noise, not real y misalignment.
Splitting median by in_table flag removes this artifact; cross-baseline
gain mean_iou 0.8535 → 0.8643 (+0.0108), PASS 12/55 → 15/55.

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
    """1D Intersection-over-Union on two y-ranges (legacy, kept for raw_iou)."""
    inter = max(0.0, min(e1, e2) - max(s1, s2))
    union = max(e1, e2) - min(s1, s2)
    if union <= 0.0:
        return 0.0
    return inter / union


def position_iou(wy: float, wh: float, oy: float, oh: float) -> float:
    """R53 (2026-04-29): position-focused IoU metric for Phase 2.

    Decouples y-position alignment from height-rendering convention.
    Phase 2 cascade work targets y-position correctness; height differences
    (Oxi grid-pitch-snapped vs Word natural font line-height) are a
    separate rendering convention issue (~0.5pt per line, ceiling at
    0.973 with strict IoU).

    Formula: 1 - |dy| / max(wh, oh), clamped to [0, 1].
        dy=0 (perfect alignment) → 1.0
        |dy|=h (1 paragraph apart) → 0.0
        |dy| > h → 0.0 (no overlap)

    Independent of relative heights; only y-start position matters.
    """
    h_max = max(wh, oh)
    if h_max <= 0:
        return 0.0
    dy_abs = abs(oy - wy)
    return max(0.0, 1.0 - dy_abs / h_max)


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
    (within same page). Returns a flat list with page info.

    Session 75 Phase D (2026-05-17): Oxi y is now LINE BOX TOP directly
    (Rust layout producer flipped). Phase C's text_y_off subtraction is
    removed — y is already in the correct convention. text_y_off remains
    in records as diagnostic only. See
    memory/session71_y_convention_refactor_design.md.
    """
    out = []
    for page_str in sorted(pages.keys(), key=int):
        page = int(page_str)
        recs = pages[page_str]
        # Sort by y (already LINE BOX TOP after Phase D)
        sorted_recs = sorted([r for r in recs if r.get("y") is not None],
                             key=lambda r: r["y"])
        for i, r in enumerate(sorted_recs):
            h = None
            for j in range(i + 1, len(sorted_recs)):
                nr = sorted_recs[j]
                if nr["y"] > r["y"]:
                    h = nr["y"] - r["y"]
                    break
            if h is None or h <= 0:
                h = DEFAULT_LINE_H
            out.append({**r, "page": page, "h": h, "y_end": r["y"] + h})
    return out


def diff_doc(doc_id: str, word: dict, oxi: dict) -> dict:
    """Compute Phase 2 element IoU for a doc.

    R52 (2026-04-29): Word COM Information(6) reports paragraph y at the
    paragraph CELL TOP (= cursor_y, before centering offset). Oxi
    dump-layout reports text element y at the RENDERED position (= cursor_y +
    centering offset within grid pitch, typically +2pt for 14pt natural
    in 18pt grid). This convention difference inflates measured dy by
    ~+2pt across all paragraphs even when pixels match.

    To measure ACTUAL misalignment (not convention noise), the tool
    computes systematic per-doc dy offset (median across matched
    same-page paragraphs) and subtracts it from Oxi y before IoU.

    The raw_mean_iou is also reported for comparison.
    """
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
    raw_matches: list[dict] = []

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
            raw_matches.append({
                "word_i": wp["i"],
                "word_page": wpage,
                "word_y": wp["y"],
                "word_h": wp["h"],
                "matched": False,
            })
            continue

        used.add(best)
        opage, oidx = best
        orec = oxi_by_page[opage][oidx][1]
        # S166 (2026-05-21): use visual text top (= line_top + text_y_off) for
        # oxi_y, to align with Word's Information(6) which returns text top.
        # OXI_IOU_USE_LINE_TOP=1 restores prior behavior (line top compare).
        use_visual = os.environ.get("OXI_IOU_USE_LINE_TOP") is None
        oxi_text_y_off = orec.get("text_y_off", 0.0) if use_visual else 0.0
        raw_matches.append({
            "word_i": wp["i"],
            "word_page": wpage,
            "word_y": wp["y"],
            "word_h": wp["h"],
            "in_table": bool(wp.get("in_table", False)),
            "oxi_page": opage,
            "oxi_y": orec["y"] + oxi_text_y_off,
            "oxi_h": orec["h"],
            "matched": True,
        })

    # R54 (2026-05-17): per-(in_table) median_dy bimodal repair.
    # See module docstring for motivation.
    same_page = [m for m in raw_matches
                 if m["matched"] and m.get("oxi_page") == m["word_page"]]
    same_page_body = [m for m in same_page if not m.get("in_table")]
    same_page_tab = [m for m in same_page if m.get("in_table")]

    def _median(xs: list[float]) -> float:
        if not xs:
            return 0.0
        xs_sorted = sorted(xs)
        return xs_sorted[len(xs_sorted) // 2]

    median_dy_body = _median([m["oxi_y"] - m["word_y"] for m in same_page_body])
    median_dy_tab = _median([m["oxi_y"] - m["word_y"] for m in same_page_tab])
    # Backward-compat doc-wide median (reported but not used for IoU)
    median_dy = _median([m["oxi_y"] - m["word_y"] for m in same_page])
    # Fallback: if one cohort is missing, use the other so we don't fall
    # back to 0.0 (which would inflate measured residual)
    if not same_page_body and same_page_tab:
        median_dy_body = median_dy_tab
    if not same_page_tab and same_page_body:
        median_dy_tab = median_dy_body

    # Build final matches with both raw and adjusted IoU
    matches = []
    for m in raw_matches:
        if not m["matched"]:
            matches.append({
                "word_i": m["word_i"],
                "word_page": m["word_page"],
                "word_y": round(m["word_y"], 2),
                "word_h": round(m["word_h"], 2),
                "oxi_y": None, "oxi_h": None,
                "iou_raw": None, "iou_adj": None, "matched": False,
            })
            continue
        wy, wh = m["word_y"], m["word_h"]
        oy, oh = m["oxi_y"], m["oxi_h"]
        offset = median_dy_tab if m.get("in_table") else median_dy_body
        # R53: position-focused IoU is the Phase 2 gate metric. Raw and
        # adjusted yrange-IoU kept for diagnosis.
        if m["oxi_page"] == m["word_page"]:
            iou_raw = yrange_iou(wy, wy + wh, oy, oy + oh)
            iou_yrange_adj = yrange_iou(wy, wy + wh, oy - offset, oy - offset + oh)
            iou_pos = position_iou(wy, wh, oy - offset, oh)
        else:
            iou_raw = 0.0
            iou_yrange_adj = 0.0
            iou_pos = 0.0
        matches.append({
            "word_i": m["word_i"],
            "word_page": m["word_page"],
            "word_y": round(wy, 2),
            "word_h": round(wh, 2),
            "in_table": m.get("in_table", False),
            "oxi_page": m["oxi_page"],
            "oxi_y": round(oy, 2),
            "oxi_h": round(oh, 2),
            "iou_raw": round(iou_raw, 4),
            "iou_yrange_adj": round(iou_yrange_adj, 4),
            "iou_pos": round(iou_pos, 4),  # Phase 2 gate metric
            "matched": True,
        })

    matched = [m for m in matches if m["matched"]]
    n_matched = len(matched)
    if n_matched > 0:
        mean_iou_raw = sum(m["iou_raw"] for m in matched) / n_matched
        mean_iou_yrange = sum(m["iou_yrange_adj"] for m in matched) / n_matched
        mean_iou_pos = sum(m["iou_pos"] for m in matched) / n_matched
        n_iou_high = sum(1 for m in matched if m["iou_pos"] >= 0.99)
        n_iou_zero = sum(1 for m in matched if m["iou_pos"] == 0.0)
    else:
        mean_iou_raw = 0.0
        mean_iou_yrange = 0.0
        mean_iou_pos = 0.0
        n_iou_high = 0
        n_iou_zero = 0

    return {
        "doc_id": doc_id,
        "n_word_paras": len(word_paras),
        "n_matched": n_matched,
        "n_unmatched": len(matches) - n_matched,
        "median_dy": round(median_dy, 2),
        "median_dy_body": round(median_dy_body, 2),
        "median_dy_table": round(median_dy_tab, 2),
        "n_same_page_body": len(same_page_body),
        "n_same_page_table": len(same_page_tab),
        # Primary gate metric: position-focused IoU (R53, 2026-04-29).
        # Decoupled from height rendering convention (Oxi grid-pitch vs
        # Word natural). Measures real y-position alignment.
        "mean_iou": round(mean_iou_pos, 4),
        "mean_iou_yrange_adj": round(mean_iou_yrange, 4),  # legacy R52 metric
        "mean_iou_raw": round(mean_iou_raw, 4),
        "n_iou_high": n_iou_high,
        "n_iou_zero": n_iou_zero,
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
