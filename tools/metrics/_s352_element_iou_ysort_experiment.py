"""
S352 — experiment: recompute element_iou for 31420af with Y-sorted Word
iteration (matching Oxi's existing derive_oxi_heights), to verify the
S349 finding that asymmetric iteration causes spurious low IoU.

This does NOT modify element_iou_diff.py. Reads the existing
pagination_word + pagination_oxi JSON, computes h with TWO methods:
- Method A (existing): Word iterates by index order
- Method B (proposed): Word iterates by Y order (matches Oxi)

Reports per-paragraph IoU delta and aggregate mean_iou delta.
"""
import json
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

MIN_MATCH_LEN = 4
DEFAULT_LINE_H = 18.0


def normalize_text(t: str) -> str:
    return re.sub(r"[\s　]+", "", t or "")


def _is_h_reference(text: str) -> bool:
    return len(normalize_text(text or "")) >= MIN_MATCH_LEN


def derive_word_heights_index_order(paragraphs):
    """Existing: iterate by index, take next h-ref same-page."""
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
            if not _is_h_reference(np.get("text", "")):
                continue
            if np["y"] > p["y"]:
                h = np["y"] - p["y"]
                break
        if h is None or h <= 0:
            h = DEFAULT_LINE_H
        out.append({**p, "h": h})
    return out


def derive_word_heights_y_sorted(paragraphs):
    """Proposed: per-page Y-sort, then next-by-Y h-ref."""
    out = []
    by_page = {}
    for p in paragraphs:
        if p.get("y") is None or p.get("page") is None:
            continue
        by_page.setdefault(p["page"], []).append(p)
    for page, recs in by_page.items():
        sorted_recs = sorted(recs, key=lambda r: r["y"])
        for i, r in enumerate(sorted_recs):
            h = None
            for j in range(i + 1, len(sorted_recs)):
                nr = sorted_recs[j]
                if not _is_h_reference(nr.get("text", "")):
                    continue
                if nr["y"] > r["y"]:
                    h = nr["y"] - r["y"]
                    break
            if h is None or h <= 0:
                h = DEFAULT_LINE_H
            out.append({**r, "h": h})
    return out


def derive_oxi_heights(pages):
    """Existing Oxi: Y-sort per page, next h-ref."""
    out = []
    for page_str in sorted(pages.keys(), key=int):
        page = int(page_str)
        recs = pages[page_str]
        sorted_recs = sorted([r for r in recs if r.get("y") is not None], key=lambda r: r["y"])
        for i, r in enumerate(sorted_recs):
            h = None
            for j in range(i + 1, len(sorted_recs)):
                nr = sorted_recs[j]
                if not _is_h_reference(nr.get("text", "")):
                    continue
                if nr["y"] > r["y"]:
                    h = nr["y"] - r["y"]
                    break
            if h is None or h <= 0:
                h = DEFAULT_LINE_H
            out.append({**r, "page": page, "h": h})
    return out


def iou_rect(wy, wh, oy, oh):
    """1D IoU on [wy, wy+wh] vs [oy, oy+oh]."""
    if wh <= 0 or oh <= 0:
        return 0.0
    w_end = wy + wh
    o_end = oy + oh
    inter = max(0.0, min(w_end, o_end) - max(wy, oy))
    union = max(w_end, o_end) - min(wy, oy)
    return inter / union if union > 0 else 0.0


def compute_iou(word, oxi, word_derive_fn):
    word_h = word_derive_fn(word.get("paragraphs", []))
    oxi_h = derive_oxi_heights(oxi.get("pages", {}))

    # Index Oxi by text+page for matching
    oxi_by_page = {}
    for r in oxi_h:
        t = normalize_text(r.get("text", ""))
        if len(t) < MIN_MATCH_LEN:
            continue
        oxi_by_page.setdefault(r["page"], []).append((t, r))

    matches = []
    used = set()
    for wp in word_h:
        if not _is_h_reference(wp.get("text", "")):
            continue
        wt = normalize_text(wp["text"])
        page = wp["page"]
        # Best match by text prefix
        best = None
        for idx, (ot, orec) in enumerate(oxi_by_page.get(page, [])):
            if idx in used:
                continue
            # Prefix match (Oxi may truncate)
            common_len = 0
            for a, b in zip(wt, ot):
                if a == b:
                    common_len += 1
                else:
                    break
            if common_len >= MIN_MATCH_LEN:
                if best is None or common_len > best[2]:
                    best = (idx, orec, common_len)
        if best is None:
            matches.append({"matched": False, "wp": wp})
            continue
        idx, orec, common_len = best
        used.add(idx)
        iou = iou_rect(wp["y"], wp["h"], orec["y"], orec["h"])
        matches.append({
            "matched": True,
            "word_y": wp["y"], "word_h": wp["h"],
            "oxi_y": orec["y"], "oxi_h": orec["h"],
            "iou": iou,
            "wp_text": wp["text"][:30],
        })
    matched = [m for m in matches if m["matched"]]
    mean_iou = sum(m["iou"] for m in matched) / max(len(matched), 1)
    return {"mean_iou": mean_iou, "n_matched": len(matched), "n_total": len(matches), "matches": matches}


def run_corpus():
    word_dir = Path("pipeline_data/pagination_word")
    oxi_dir = Path("pipeline_data/pagination_oxi")
    docs = sorted([f.stem for f in word_dir.glob("*.json")])
    print(f"Running on {len(docs)} docs")
    print(f'{"doc":>14} {"mean_idx":>9} {"mean_y":>8} {"delta":>8} {"sign":>5}')
    pos = neg = neu = 0
    sum_a = sum_b = 0.0
    deltas = []
    for d in docs:
        wp = word_dir / f"{d}.json"
        op = oxi_dir / f"{d}.json"
        if not op.exists():
            continue
        with open(wp, encoding="utf-8") as f:
            w = json.load(f)
        with open(op, encoding="utf-8") as f:
            o = json.load(f)
        ra = compute_iou(w, o, derive_word_heights_index_order)
        rb = compute_iou(w, o, derive_word_heights_y_sorted)
        delta = rb["mean_iou"] - ra["mean_iou"]
        sum_a += ra["mean_iou"]
        sum_b += rb["mean_iou"]
        sign = "+" if delta > 0.0005 else ("-" if delta < -0.0005 else "0")
        if sign == "+":
            pos += 1
        elif sign == "-":
            neg += 1
        else:
            neu += 1
        deltas.append((delta, d, ra["mean_iou"], rb["mean_iou"]))
        if abs(delta) >= 0.005:
            print(f'{d:>14} {ra["mean_iou"]:>9.4f} {rb["mean_iou"]:>8.4f} {delta:>+8.4f} {sign:>5}')
    print()
    print(f"Corpus: pos={pos} neg={neg} neutral={neu} total={pos+neg+neu}")
    print(f"Mean A (INDEX) = {sum_a/(pos+neg+neu):.4f}")
    print(f"Mean B (Y-SORT) = {sum_b/(pos+neg+neu):.4f}")
    print(f"Net delta = {(sum_b - sum_a) / (pos+neg+neu):+.4f}")
    deltas.sort()
    print("\nTop 10 worst delta:")
    for d, doc, a, b in deltas[:10]:
        print(f"  Δ={d:+.4f}  {doc}  A={a:.4f}  B={b:.4f}")
    print("\nTop 10 best delta:")
    for d, doc, a, b in deltas[-10:]:
        print(f"  Δ={d:+.4f}  {doc}  A={a:.4f}  B={b:.4f}")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "corpus":
        run_corpus()
        sys.exit(0)
    doc = "31420af1a08f"
    with open(f"pipeline_data/pagination_word/{doc}.json", encoding="utf-8") as f:
        word = json.load(f)
    with open(f"pipeline_data/pagination_oxi/{doc}.json", encoding="utf-8") as f:
        oxi = json.load(f)

    res_idx = compute_iou(word, oxi, derive_word_heights_index_order)
    res_y = compute_iou(word, oxi, derive_word_heights_y_sorted)

    print(f"=== {doc} ===")
    print(f"Method A (Word INDEX order, existing): mean_iou={res_idx['mean_iou']:.4f}, matched={res_idx['n_matched']}/{res_idx['n_total']}")
    print(f"Method B (Word Y-SORTED, proposed):   mean_iou={res_y['mean_iou']:.4f}, matched={res_y['n_matched']}/{res_y['n_total']}")
    print(f"Delta: {res_y['mean_iou'] - res_idx['mean_iou']:+.4f}")
    print()

    # Show top paragraph-level IoU changes
    # Match by (word_y, oxi_y) tuple
    by_key_a = {(m.get("word_y"), m.get("oxi_y")): m for m in res_idx["matches"] if m["matched"]}
    by_key_b = {(m.get("word_y"), m.get("oxi_y")): m for m in res_y["matches"] if m["matched"]}
    deltas = []
    for k, m_a in by_key_a.items():
        m_b = by_key_b.get(k)
        if m_b is None:
            continue
        d = m_b["iou"] - m_a["iou"]
        if abs(d) >= 0.001:
            deltas.append((d, k, m_a, m_b))
    deltas.sort()
    print("Top 15 worst paragraph IoU changes (improvements at bottom):")
    for d, k, m_a, m_b in deltas[:7] + deltas[-7:]:
        print(f"  Δ={d:+.4f}  wy={m_a['word_y']:.1f} wh_idx={m_a['word_h']:.1f} wh_y={m_b['word_h']:.1f} oy={m_a['oxi_y']:.1f} oh={m_a['oxi_h']:.1f}  '{m_a['wp_text']}'")
