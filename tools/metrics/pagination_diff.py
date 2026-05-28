"""Cross-join Word per-paragraph page index with Oxi per-page paragraph
records, compute the Phase 1 pagination gate for each baseline doc.

Phase 1 gate (2026-04-28 redesigned methodology):
- Per doc, match each Word paragraph to an Oxi paragraph by text prefix
  (anchored to the same expected page, falling back to nearby pages if
  the cascade has shifted text across boundary).
- For each matched pair, page_delta = oxi_page - word_page.
- Per-doc pass = all matched paragraphs have page_delta == 0.
- Per-doc continuous score = fraction of matched paragraphs with delta=0.
- Aggregate: pass_rate (binary) and mean_score (continuous).

The continuous score is the diagnostic signal — pass/fail is the gate.

Output: pipeline_data/pagination_diff/<doc_id>.json (per-doc detail)
        pipeline_data/pagination_diff/_summary.json (cross-doc gate state)

Run from repo root:
    python tools/metrics/pagination_diff.py            # all docs with both inputs
    python tools/metrics/pagination_diff.py 2ea81a
"""
from __future__ import annotations

import json
import os
import re
import sys
from collections import Counter

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
WORD_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_word")
OXI_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_oxi")
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_diff")

# Min text length to consider a paragraph matchable. Empty / whitespace-only
# paragraphs are unmatchable by text prefix and are excluded from the gate
# (they affect Y but the gate is page-membership, not Y).
MIN_MATCH_LEN = 2

# Page search window when text doesn't match on the expected page —
# allows for cascade-shifted text. Wider = more lenient matching but
# more chance of false matches across structurally similar paragraphs.
PAGE_SEARCH_RADIUS = 3


def normalize_text(s: str) -> str:
    if not s:
        return ""
    s = s.replace("　", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def load_word(doc_id: str) -> dict | None:
    path = os.path.join(WORD_DIR, f"{doc_id}.json")
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_oxi(doc_id: str) -> dict | None:
    path = os.path.join(OXI_DIR, f"{doc_id}.json")
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def diff_doc(doc_id: str, word: dict, oxi: dict) -> dict:
    # Build (page → list of (text_norm, record)) for Oxi
    oxi_by_page: dict[int, list[tuple[str, dict]]] = {}
    for page_str, recs in oxi["pages"].items():
        page = int(page_str)
        for r in recs:
            t = normalize_text(r["text"])
            if len(t) < MIN_MATCH_LEN:
                continue
            oxi_by_page.setdefault(page, []).append((t, r))

    # Per-record use counter. Oxi records may concatenate text from
    # multiple cell-paragraphs that Word reports as separate paragraphs
    # (e.g. a 4-cell table row "千円|千円|千円|千円" becomes one Oxi entry
    # "千円千円千円千円"). For each match attempt, the record's capacity is
    # max(1, count(prefix in record_text)). This allows the matcher to map
    # multiple Word paragraphs to the same Oxi entry when the entry literally
    # contains that many repetitions of the matched prefix — closing the
    # R7.32 cell_paragraph_index gap for docs where Oxi joins per-cell text.
    #
    # S400 (2026-05-28) CORRECTION to S399: bare text-prefix matching CAN
    # cause spurious deltas on repetitive-text tables in principle, but
    # ed025's specific Phase 1 -1 delta is NOT a matcher artifact — it
    # is a REAL Oxi layout difference. Per-page (y,x)-unique × × × counts
    # for ed025: pages 5/6/9/10/15 match Word exactly; page 13 Word=28
    # Oxi=29 (+1), page 14 Word=39 Oxi=38 (-1) → 1 cell that Word places
    # on page 14 is on Oxi page 13. The matcher correctly identifies
    # this. A column-aware tie-break (x distance + in_table cohort) was
    # tested in S400 and produced no corpus change; reverted to keep
    # this tool simple. Engine-side fix for ed025's misplaced cell
    # remains an open per-doc target.
    used: dict[tuple[int, int], int] = {}

    matches: list[dict] = []
    word_paras = word.get("paragraphs", [])
    for wp in word_paras:
        wt = normalize_text(wp["text"])
        if len(wt) < MIN_MATCH_LEN:
            continue
        wpage = wp.get("page")
        if wpage is None:
            continue

        best = None
        best_dist = None
        # Search expected page first, then expand outward
        for radius in range(0, PAGE_SEARCH_RADIUS + 1):
            for sign in ((0,) if radius == 0 else (-1, +1)):
                cand_page = wpage + sign * radius
                if cand_page < 1:
                    continue
                cand_recs = oxi_by_page.get(cand_page, [])
                for idx, (ot, _orec) in enumerate(cand_recs):
                    # Match: prefix-equality up to the shorter of the two,
                    # min length MIN_MATCH_LEN.
                    n = min(len(wt), len(ot))
                    if n < MIN_MATCH_LEN:
                        continue
                    if wt[:n] != ot[:n]:
                        continue
                    # Capacity: how many times the matched prefix appears
                    # in the Oxi record's text (≥1 by construction).
                    capacity = max(1, ot.count(wt[:n]))
                    if used.get((cand_page, idx), 0) >= capacity:
                        continue
                    # Distance metric for tie-breaking: pick exact-page
                    # match over nearby-page match.
                    dist = (radius, -n)
                    if best is None or dist < best_dist:
                        best = (cand_page, idx)
                        best_dist = dist
            if best is not None and radius == 0:
                # Same-page match found; don't search further pages.
                break

        if best is None:
            matches.append({
                "word_i": wp["i"],
                "word_page": wpage,
                "oxi_page": None,
                "page_delta": None,
                "text": wt[:30],
                "matched": False,
            })
            continue
        used[best] = used.get(best, 0) + 1
        opage, oidx = best
        matches.append({
            "word_i": wp["i"],
            "word_page": wpage,
            "oxi_page": opage,
            "page_delta": opage - wpage,
            "text": wt[:30],
            "matched": True,
        })

    matched = [m for m in matches if m["matched"]]
    unmatched = [m for m in matches if not m["matched"]]
    n_matched = len(matched)
    n_zero = sum(1 for m in matched if m["page_delta"] == 0)
    n_pos = sum(1 for m in matched if m["page_delta"] is not None and m["page_delta"] > 0)
    n_neg = sum(1 for m in matched if m["page_delta"] is not None and m["page_delta"] < 0)
    delta_hist = Counter(m["page_delta"] for m in matched if m["page_delta"] is not None)

    score = (n_zero / n_matched) if n_matched else 0.0
    # Binary pass: every matched paragraph on its expected page.
    # Unmatched paragraphs are NOT failures (text may legitimately differ
    # due to font substitution / encoding) but they reduce confidence;
    # surface in the summary.
    pass_binary = (n_matched > 0) and (n_zero == n_matched)

    return {
        "doc_id": doc_id,
        "word_filename": word.get("filename"),
        "word_n_pages": word.get("n_pages"),
        "oxi_n_pages": oxi.get("n_pages"),
        "page_count_delta": (oxi.get("n_pages") or 0) - (word.get("n_pages") or 0),
        "n_word_paras": len(word_paras),
        "n_matched": n_matched,
        "n_unmatched": len(unmatched),
        "n_page_zero": n_zero,
        "n_page_positive": n_pos,
        "n_page_negative": n_neg,
        "score": round(score, 4),
        "pass": pass_binary,
        "delta_histogram": dict(sorted(delta_hist.items())),
        "matches": matches,
    }


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)

    prefix = sys.argv[1] if len(sys.argv) > 1 else None
    word_files = sorted(f for f in os.listdir(WORD_DIR) if f.endswith(".json") and not f.startswith("_"))
    doc_ids = [os.path.splitext(f)[0] for f in word_files]
    if prefix:
        doc_ids = [d for d in doc_ids if d.startswith(prefix)]
    if not doc_ids:
        print(f"no docs matched (WORD_DIR={WORD_DIR}, prefix={prefix})", file=sys.stderr)
        return 2

    summary = []
    skipped = []
    for doc_id in doc_ids:
        word = load_word(doc_id)
        oxi = load_oxi(doc_id)
        if word is None or oxi is None:
            skipped.append({"doc_id": doc_id, "has_word": word is not None, "has_oxi": oxi is not None})
            continue
        result = diff_doc(doc_id, word, oxi)
        out_path = os.path.join(OUT_DIR, f"{doc_id}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        marker = "PASS" if result["pass"] else "FAIL"
        print(f"  [{marker}] {doc_id}: score={result['score']} matched={result['n_matched']}/{result['n_matched']+result['n_unmatched']} delta_hist={result['delta_histogram']}")
        summary.append({
            "doc_id": doc_id,
            "pass": result["pass"],
            "score": result["score"],
            "n_matched": result["n_matched"],
            "n_unmatched": result["n_unmatched"],
            "page_count_delta": result["page_count_delta"],
            "delta_histogram": result["delta_histogram"],
        })

    n_pass = sum(1 for s in summary if s["pass"])
    n_total = len(summary)
    pass_rate = (n_pass / n_total) if n_total else 0.0
    mean_score = (sum(s["score"] for s in summary) / n_total) if n_total else 0.0
    summary_obj = {
        "n_total": n_total,
        "n_pass": n_pass,
        "n_fail": n_total - n_pass,
        "pass_rate": round(pass_rate, 4),
        "mean_score": round(mean_score, 4),
        "n_skipped": len(skipped),
        "docs": summary,
        "skipped": skipped,
    }
    summary_path = os.path.join(OUT_DIR, "_summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary_obj, f, ensure_ascii=False, indent=2)
    print(f"\nsummary -> {summary_path}")
    print(f"  PHASE 1 GATE: pass_rate={pass_rate:.2%} ({n_pass}/{n_total}), mean_score={mean_score:.4f}")
    if skipped:
        print(f"  skipped {len(skipped)} docs (missing Word or Oxi input)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
