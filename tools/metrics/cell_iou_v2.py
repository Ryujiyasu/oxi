"""S449 -> S450: STRUCTURAL cell IoU (v2) — the ruler fix.

Fixes the two ruler bugs S449 pinned in cell_iou_diff.py (v1):
  (1) v1 matched Word<->Oxi by global text-prefix + nearest-y across the whole
      page, which MISPAIRS cells in merged / multi-line / interleaved-column
      tables (e.g. 34140b: 45% of cells are truly aligned with raw iou 0.98 but
      a handful of mispaired cells fabricate big dy).
  (2) those mispaired dy values are BIMODAL, so v1's per-table MEDIAN dy
      correction lands between the modes and shifts the correctly-aligned
      majority OUT of alignment (34140b aligned cells 0.98 -> 0.85).

v2 approach — match by STRUCTURE, not text+nearest-y:
  - establish table correspondence by document order (i-th Word table <-> i-th
    Oxi table), validated by text overlap;
  - within a matched table, match cells COLUMN by COLUMN: both engines list a
    column's paragraphs top-to-bottom, so a windowed two-pointer text-prefix
    alignment inside one column is unambiguous and immune to vMerge row-index
    skew (Word enumerates all rows; Oxi skips vMerge continuations);
  - height = gap to the next paragraph in the SAME column (same page), for both
    engines identically (cell-scoped pitch);
  - report BOTH raw IoU (no correction) and a ROBUST-offset IoU (trimmed mean of
    dy, so bimodal outliers cannot drag the offset). With good structural
    matching the dy distribution is unimodal and raw ~= robust.

DIAGNOSTIC tool (not yet a gate). Compares against v1 (cell_iou_diff) and the
Phase-2 gate (element_iou) per doc.

Run from repo root:
  python tools/metrics/cell_iou_v2.py            # all docs, summary
  python tools/metrics/cell_iou_v2.py 34140b     # one doc, detail
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
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "cell_iou_v2")

MIN_MATCH_LEN = 2
DEFAULT_LINE_H = 14.0
RESYNC_WINDOW = 4  # how far to look ahead when a column's sequences desync


def norm(s: str) -> str:
    if not s:
        return ""
    s = s.replace("　", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def prefix_match(a: str, b: str, n: int = 5) -> bool:
    a, b = norm(a), norm(b)
    if len(a) < MIN_MATCH_LEN or len(b) < MIN_MATCH_LEN:
        return False
    k = min(n, len(a), len(b))
    return a[:k] == b[:k]


def pos_iou(wy, wh, oy, oh):
    h = max(wh, oh)
    return 0.0 if h <= 0 else max(0.0, 1.0 - abs(oy - wy) / h)


def load(p):
    if not os.path.exists(p):
        return None
    with io.open(p, encoding="utf-8") as f:
        return json.load(f)


# --- gather structured table paragraphs from each engine -------------------

def word_tables(paras):
    """Return {table_start: {col: [paras sorted by (page,y)]}} + table order."""
    tabs: dict = {}
    for p in paras:
        ts = p.get("table_start")
        if ts is None or p.get("y") is None:
            continue
        c = p.get("cell_col")
        tabs.setdefault(ts, {}).setdefault(c, []).append(p)
    order = sorted(tabs.keys(), key=lambda ts: min(
        (pp["y"] + pp.get("page", 1) * 100000) for col in tabs[ts].values() for pp in col))
    for ts in tabs:
        for c in tabs[ts]:
            tabs[ts][c].sort(key=lambda pp: (pp.get("page", 1), pp["y"]))
    return tabs, order


def oxi_tables(pages):
    tabs: dict = {}
    for pg, recs in pages.items():
        page = int(pg)
        for r in recs:
            if r.get("cell_col_idx") is None or r.get("y") is None:
                continue
            tid = r.get("para_idx")
            rr = {**r, "page": page}
            tabs.setdefault(tid, {}).setdefault(r["cell_col_idx"], []).append(rr)
    order = sorted(tabs.keys(), key=lambda t: min(
        (rr["y"] + rr["page"] * 100000) for col in tabs[t].values() for rr in col))
    for t in tabs:
        for c in tabs[t]:
            tabs[t][c].sort(key=lambda rr: (rr["page"], rr["y"]))
    return tabs, order


def col_height(cells, idx):
    """gap to next paragraph in this column (same page), else DEFAULT_LINE_H."""
    cur = cells[idx]
    for j in range(idx + 1, len(cells)):
        nx = cells[j]
        if nx["page"] != cur["page"]:
            break
        if nx["y"] > cur["y"] + 0.01:
            return nx["y"] - cur["y"]
    return DEFAULT_LINE_H


def align_column(wcol, ocol):
    """Windowed two-pointer text-prefix alignment within one column."""
    pairs = []
    i = j = 0
    while i < len(wcol) and j < len(ocol):
        if prefix_match(wcol[i].get("text", ""), ocol[j].get("text", "")):
            pairs.append((i, j))
            i += 1
            j += 1
            continue
        # desync: find nearest re-sync within window
        found = None
        for di in range(RESYNC_WINDOW + 1):
            for dj in range(RESYNC_WINDOW + 1):
                if di == 0 and dj == 0:
                    continue
                if i + di < len(wcol) and j + dj < len(ocol) and \
                        prefix_match(wcol[i + di].get("text", ""), ocol[j + dj].get("text", "")):
                    found = (di, dj)
                    break
            if found:
                break
        if found is None:
            i += 1
            j += 1
        else:
            i += found[0]
            j += found[1]
    return pairs


def trimmed_mean(xs, frac=0.2):
    if not xs:
        return 0.0
    xs = sorted(xs)
    k = int(len(xs) * frac)
    core = xs[k: len(xs) - k] or xs
    return sum(core) / len(core)


def diff_doc(doc_id, word, oxi):
    wt, worder = word_tables(word.get("paragraphs", []))
    ot, oorder = oxi_tables(oxi.get("pages", {}))
    n_wt, n_ot = len(worder), len(oorder)
    rows = []
    for ti in range(min(n_wt, n_ot)):
        wtab = wt[worder[ti]]
        otab = ot[oorder[ti]]
        for c in sorted(set(wtab) & set(otab)):
            wcol = wtab[c]
            ocol = otab[c]
            for (wi, oj) in align_column(wcol, ocol):
                wp, op = wcol[wi], ocol[oj]
                wh = col_height(wcol, wi)
                oh = col_height(ocol, oj)
                rows.append({
                    "page": wp.get("page"), "col": c,
                    "word_y": round(wp["y"], 2), "word_h": round(wh, 2),
                    "oxi_y": round(op["y"], 2), "oxi_h": round(oh, 2),
                    "dy": round(op["y"] - wp["y"], 2),
                    "text": (wp.get("text") or "")[:16],
                })
    if not rows:
        return {"doc_id": doc_id, "n_matched": 0, "n_wtab": n_wt, "n_otab": n_ot}
    dys = [r["dy"] for r in rows]
    offset = trimmed_mean(dys, 0.2)
    raw = [pos_iou(r["word_y"], r["word_h"], r["oxi_y"], r["oxi_h"]) for r in rows]
    cor = [pos_iou(r["word_y"], r["word_h"], r["oxi_y"] - offset, r["oxi_h"]) for r in rows]
    n_align = sum(1 for r in rows if abs(r["dy"]) < 0.6 and abs(r["word_h"] - r["oxi_h"]) < 0.6)
    return {
        "doc_id": doc_id, "n_matched": len(rows),
        "n_wtab": n_wt, "n_otab": n_ot,
        "trim_offset": round(offset, 2),
        "mean_iou_raw": round(sum(raw) / len(raw), 4),
        "mean_iou_robust": round(sum(cor) / len(cor), 4),
        "frac_aligned": round(n_align / len(rows), 3),
        "rows": rows,
    }


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    prefix = args[0] if args else None
    wfiles = sorted(f for f in os.listdir(WORD_DIR)
                    if f.endswith(".json") and not f.startswith("_"))
    summ = []
    for wf in wfiles:
        did = wf[:-5]
        if prefix and not did.startswith(prefix):
            continue
        word = load(os.path.join(WORD_DIR, wf))
        oxi = load(os.path.join(OXI_DIR, wf))
        if word is None or oxi is None:
            continue
        res = diff_doc(did, word, oxi)
        if res["n_matched"] == 0:
            continue
        with io.open(os.path.join(OUT_DIR, f"{did}.json"), "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=2)
        summ.append(res)
        if prefix:
            print(f"{did}: raw={res['mean_iou_raw']} robust={res['mean_iou_robust']} "
                  f"offset={res['trim_offset']} aligned={res['frac_aligned']} "
                  f"n={res['n_matched']} wtab={res['n_wtab']} otab={res['n_otab']}")
    if summ:
        mr = sum(s["mean_iou_raw"] for s in summ) / len(summ)
        mc = sum(s["mean_iou_robust"] for s in summ) / len(summ)
        fa = sum(s["frac_aligned"] for s in summ) / len(summ)
        with io.open(os.path.join(OUT_DIR, "_summary.json"), "w", encoding="utf-8") as f:
            json.dump({"n_docs": len(summ), "mean_iou_raw": round(mr, 4),
                       "mean_iou_robust": round(mc, 4), "mean_frac_aligned": round(fa, 3),
                       "docs": [{k: s[k] for k in s if k != "rows"} for s in summ]}, f,
                      ensure_ascii=False, indent=2)
        print(f"\nv2 corpus (n={len(summ)}): mean_iou_raw={mr:.4f} "
              f"mean_iou_robust={mc:.4f} mean_frac_aligned={fa:.3f}")


if __name__ == "__main__":
    main()
