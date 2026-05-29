"""S432 (2026-05-29): cell-aware element IoU — DIAGNOSTIC (not a gate).

Motivation (S431): the Phase-2 element_iou_diff.py derives each paragraph's
height as `next_paragraph_y - this_y`. Inside tables this is WRONG: Word COM
enumerates cells row-major but Information(6) y is non-monotonic across a
row, so a single-paragraph header cell (e.g. tokumei 「名称」) gets its
height computed as the gap to a far-away paragraph (word_h≈84.5 for a ~22pt
cell), fabricating a near-zero IoU even though Oxi renders the cell
correctly. This produced the spurious convergent −81.6pt "bug" across the
tokumei_08_01 family (a1d6e4/6514f2/d4d126/de6e32).

Fix: derive heights **cell-scoped**. A paragraph's height = gap to the next
paragraph IN THE SAME CELL (same table + row + column); if it is the only /
last paragraph in its cell, fall back to DEFAULT_LINE_H. This removes the
cross-cell jump entirely, so position_iou measures true y-alignment.

Inputs (both already carry structural cell coords as of S432):
  - pipeline_data/pagination_word/<doc>.json : paragraphs[] with
      cell_row, cell_col, table_start  (added in measure_pagination_word.py)
  - pipeline_data/pagination_oxi/<doc>.json  : pages{} recs with
      para_idx (table id), cell_row_idx, cell_col_idx, cell_para_idx

This is a DIAGNOSTIC tool (S417 lesson: instrument before re-gating). It
does NOT replace element_iou_diff.py / the Phase-2 gate. It reports, per
doc, the cell-scoped mean position_iou and the cells whose verdict flips
versus the global-next-y metric (i.e. the S431 artifacts it repairs).

Run from repo root:
  python tools/metrics/cell_iou_diff.py            # all docs with both inputs
  python tools/metrics/cell_iou_diff.py a1d6e4     # prefix filter
  python tools/metrics/cell_iou_diff.py --verbose a1d6e4
"""
from __future__ import annotations

import io
import json
import os
import re
import sys

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
WORD_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_word")
OXI_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_oxi")
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "cell_iou_diff")

MIN_MATCH_LEN = 2
DEFAULT_LINE_H = 14.0
PAGE_SEARCH_RADIUS = 1


def normalize_text(s: str) -> str:
    if not s:
        return ""
    s = s.replace("　", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def position_iou(wy: float, wh: float, oy: float, oh: float) -> float:
    """1 - |dy| / max(wh, oh), clamped [0,1] (same formula as element_iou)."""
    h_max = max(wh, oh)
    if h_max <= 0:
        return 0.0
    return max(0.0, 1.0 - abs(oy - wy) / h_max)


def load(path: str):
    if not os.path.exists(path):
        return None
    with io.open(path, encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Height derivation
# ---------------------------------------------------------------------------

def _word_cell_key(r: dict):
    if r.get("table_start") is None:
        return None
    return (r["table_start"], r.get("cell_row"), r.get("cell_col"))


def derive_word(paragraphs: list[dict]) -> list[dict]:
    """Attach cell-scoped height to each Word paragraph.

    Body paragraphs (no table_start): height = gap to next body paragraph on
    the same page (legacy behavior — body y IS monotonic, so this is fine).
    Table paragraphs: height = gap to next paragraph in the SAME cell;
    fallback DEFAULT_LINE_H.
    """
    out = []
    # cell-scoped: index paragraphs by cell key, in i-order
    by_cell: dict = {}
    for r in paragraphs:
        k = _word_cell_key(r)
        if k is not None:
            by_cell.setdefault(k, []).append(r)

    for idx, p in enumerate(paragraphs):
        if p.get("y") is None or p.get("page") is None:
            continue
        k = _word_cell_key(p)
        h = None
        if k is None:
            # body: next body paragraph same page with larger y
            for j in range(idx + 1, len(paragraphs)):
                np = paragraphs[j]
                if np.get("page") != p["page"]:
                    break
                if np.get("y") is None or _word_cell_key(np) is not None:
                    continue
                if len(normalize_text(np.get("text", ""))) < MIN_MATCH_LEN:
                    continue
                if np["y"] > p["y"]:
                    h = np["y"] - p["y"]
                    break
        else:
            # table: next paragraph in same cell with larger y
            sibs = by_cell.get(k, [])
            cand = [s["y"] for s in sibs
                    if s.get("y") is not None and s.get("page") == p["page"]
                    and s["y"] > p["y"] + 0.01]
            if cand:
                h = min(cand) - p["y"]
        if h is None or h <= 0:
            h = DEFAULT_LINE_H
        out.append({**p, "h": h, "y_end": p["y"] + h})
    return out


def _oxi_cell_key(r: dict):
    if r.get("cell_col_idx") is None:
        return None
    return (r.get("para_idx"), r.get("cell_row_idx"), r.get("cell_col_idx"))


def derive_oxi(pages: dict) -> list[dict]:
    """Attach cell-scoped height to each Oxi record (flattened with page)."""
    flat = []
    for page_str, recs in pages.items():
        page = int(page_str)
        for r in recs:
            if r.get("y") is None:
                continue
            flat.append({**r, "page": page})

    by_cell: dict = {}
    body = []
    for r in flat:
        k = _oxi_cell_key(r)
        if k is not None:
            by_cell.setdefault(k, []).append(r)
        else:
            body.append(r)
    body_sorted = sorted(body, key=lambda r: (r["page"], r["y"]))

    out = []
    for r in flat:
        k = _oxi_cell_key(r)
        h = None
        if k is None:
            # body: next body rec same page larger y
            same = [b for b in body_sorted if b["page"] == r["page"] and b["y"] > r["y"] + 0.01
                    and len(normalize_text(b.get("text", ""))) >= MIN_MATCH_LEN]
            if same:
                h = same[0]["y"] - r["y"]
        else:
            cand = [s["y"] for s in by_cell.get(k, [])
                    if s["page"] == r["page"] and s["y"] > r["y"] + 0.01]
            if cand:
                h = min(cand) - r["y"]
        if h is None or h <= 0:
            h = DEFAULT_LINE_H
        out.append({**r, "h": h, "y_end": r["y"] + h})
    return out


# ---------------------------------------------------------------------------
# Matching (text-prefix + same page + nearest y — same as element_iou)
# ---------------------------------------------------------------------------

def match(word_paras: list[dict], oxi_flat: list[dict]):
    # bucket oxi by page
    by_page: dict = {}
    for r in oxi_flat:
        if len(normalize_text(r.get("text", ""))) < MIN_MATCH_LEN:
            continue
        by_page.setdefault(r["page"], []).append(r)

    used = set()
    matches = []
    for wp in word_paras:
        wt = normalize_text(wp.get("text", ""))
        if len(wt) < MIN_MATCH_LEN:
            continue
        best = None
        best_dy = None
        for dp in range(-PAGE_SEARCH_RADIUS, PAGE_SEARCH_RADIUS + 1):
            for k, op in enumerate(by_page.get(wp["page"] + dp, [])):
                key = (wp["page"] + dp, k)
                if key in used:
                    continue
                ot = normalize_text(op.get("text", ""))
                if not (ot.startswith(wt[:6]) or wt.startswith(ot[:6])):
                    continue
                dy = abs(op["y"] - wp["y"])
                if best is None or dy < best_dy:
                    best = (key, op)
                    best_dy = dy
        if best is not None:
            used.add(best[0])
            op = best[1]
            matches.append((wp, op))
    return matches


def diff_doc(doc_id: str, word: dict, oxi: dict, verbose: bool = False) -> dict:
    word_paras = derive_word(word.get("paragraphs", []))
    oxi_flat = derive_oxi(oxi.get("pages", {}))
    pairs = match(word_paras, oxi_flat)

    # per-(in_table) median dy correction (mirrors element_iou R54)
    def median(xs):
        xs = sorted(xs)
        n = len(xs)
        return 0.0 if n == 0 else (xs[n // 2] if n % 2 else (xs[n // 2 - 1] + xs[n // 2]) / 2)

    dys_tbl = [op["y"] - wp["y"] for wp, op in pairs if wp.get("table_start") is not None]
    dys_body = [op["y"] - wp["y"] for wp, op in pairs if wp.get("table_start") is None]
    med_tbl = median(dys_tbl)
    med_body = median(dys_body)

    rows = []
    for wp, op in pairs:
        in_tbl = wp.get("table_start") is not None
        med = med_tbl if in_tbl else med_body
        oy_adj = op["y"] - med
        iou = position_iou(wp["y"], wp["h"], oy_adj, op["h"])
        iou_raw = position_iou(wp["y"], wp["h"], op["y"], op["h"])
        rows.append({
            "word_i": wp.get("i"), "page": wp.get("page"), "in_table": in_tbl,
            "word_y": round(wp["y"], 2), "word_h": round(wp["h"], 2),
            "oxi_y": round(op["y"], 2), "oxi_h": round(op["h"], 2),
            "iou": round(iou, 4), "iou_raw": round(iou_raw, 4),
            "text": (wp.get("text") or "")[:16],
        })

    ious = [r["iou"] for r in rows]
    mean_iou = sum(ious) / len(ious) if ious else 0.0
    n_high = sum(1 for v in ious if v >= 0.99)
    return {
        "doc_id": doc_id,
        "n_matched": len(rows),
        "median_dy_table": round(med_tbl, 2),
        "median_dy_body": round(med_body, 2),
        "mean_iou": round(mean_iou, 4),
        "n_iou_high": n_high,
        "frac_iou_high": round(n_high / len(rows), 4) if rows else 0.0,
        "matches": rows,
    }


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    verbose = "--verbose" in sys.argv
    prefix = args[0] if args else None

    word_files = sorted(f for f in os.listdir(WORD_DIR) if f.endswith(".json") and not f.startswith("_"))
    summary = []
    for wf in word_files:
        doc_id = wf[:-5]
        if prefix and not doc_id.startswith(prefix):
            continue
        word = load(os.path.join(WORD_DIR, wf))
        oxi = load(os.path.join(OXI_DIR, wf))
        if word is None or oxi is None:
            continue
        # skip docs whose Word data predates S432 (no cell coords) — detect by
        # presence of the new key on any table paragraph.
        has_cellcoords = any("table_start" in p for p in word.get("paragraphs", []))
        res = diff_doc(doc_id, word, oxi, verbose)
        res["cellcoords"] = has_cellcoords
        with io.open(os.path.join(OUT_DIR, f"{doc_id}.json"), "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=2)
        summary.append({"doc_id": doc_id, "mean_iou": res["mean_iou"],
                        "n_matched": res["n_matched"], "frac_iou_high": res["frac_iou_high"],
                        "cellcoords": has_cellcoords})
        tag = "" if has_cellcoords else "  [NO-CELLCOORDS: rerun measure_pagination_word]"
        print(f"{doc_id}: mean_iou={res['mean_iou']:.4f} matched={res['n_matched']} "
              f"frac_high={res['frac_iou_high']:.3f}{tag}")
        if verbose:
            for r in sorted(res["matches"], key=lambda x: x["iou"])[:12]:
                print(f"    iou={r['iou']:.3f} tbl={int(r['in_table'])} "
                      f"wy={r['word_y']} wh={r['word_h']} oy={r['oxi_y']} oh={r['oxi_h']} {r['text']!r}")

    if summary:
        cc = [s for s in summary if s["cellcoords"]]
        mean = sum(s["mean_iou"] for s in cc) / len(cc) if cc else 0.0
        with io.open(os.path.join(OUT_DIR, "_summary.json"), "w", encoding="utf-8") as f:
            json.dump({"n_docs": len(summary), "n_cellcoords": len(cc),
                       "mean_iou_cellcoords": round(mean, 4), "docs": summary}, f,
                      ensure_ascii=False, indent=2)
        print(f"\nmean_iou (cellcoords docs only, n={len(cc)}): {mean:.4f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
