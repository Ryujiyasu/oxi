"""Positional matcher — match Word's Document.Paragraphs(i) enumeration
against Oxi's document-order enumeration of distinct
(para_idx, cell_row_idx, cell_col_idx, cell_para_idx) tuples.

Why: the default text-prefix matcher in `pagination_diff.py` fails on
docs with many repeated short paragraph prefixes (□, ○, "別紙のとおり",
empty cells). On d4d126, that matcher reports score=1.0 even though
actual content drifts ~-4pt/row. See `docs/spec/d4d126_drift_root_cause_s176.md`
and `memory/session176_d4d126_phase_a_correction.md` for the falsified
S173 hypothesis that motivated this rewrite.

Approach:
1. Enumerate Word paragraphs in `i` order from
   `pipeline_data/pagination_word/<doc>.json`.
2. Enumerate Oxi paragraphs in document order from
   `pipeline_data/pagination_oxi/<doc>.json`. Document order = (page,
   y, x) within each page; tuples = (para_idx, cell_row_idx, cell_col_idx,
   cell_para_idx). Each tuple's first occurrence assigns its enum_i.
3. Match Word.i <-> Oxi.enum_i positionally (1:1 by index).
4. Verify with text-prefix on each pair as a sanity check; report
   mismatches as anomalies.

Output: `pipeline_data/pagination_diff_positional/<doc>.json`,
        `pipeline_data/pagination_diff_positional/_summary.json`.

Does NOT modify the Phase 1 gate — `pagination_diff.py` is untouched
and its output `pipeline_data/pagination_diff/_summary.json` remains
authoritative. This script is a parallel diagnostic.

Run from repo root:
    python tools/metrics/pagination_diff_positional.py            # all docs
    python tools/metrics/pagination_diff_positional.py d4d126     # prefix filter
    python tools/metrics/pagination_diff_positional.py --verbose d4d126
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
OUT_DIR = os.path.join(REPO_ROOT, "pipeline_data", "pagination_diff_positional")


def normalize_text(s: str) -> str:
    if not s:
        return ""
    s = s.replace("　", " ").replace("　", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def load_json(path: str) -> dict | None:
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def enumerate_oxi(oxi: dict) -> list[dict]:
    """Walk Oxi pages in (page, y, x) order, assign enum_i to each distinct
    (para_idx, cell_row_idx, cell_col_idx, cell_para_idx) tuple at its first
    occurrence. Returns ordered list of {enum_i, page, y, x, text, key}.
    """
    seen: dict[tuple, int] = {}
    enum: list[dict] = []
    # Pages may not be in sorted order in the dict; sort by int.
    page_keys = sorted(oxi["pages"].keys(), key=lambda k: int(k))
    for pk in page_keys:
        recs = oxi["pages"][pk]
        # Records within a page are already sorted by (y, x) at write time
        # (measure_pagination_oxi.py line ~135), but re-sort defensively.
        recs_sorted = sorted(recs, key=lambda r: (r.get("y", 0.0), r.get("x", 0.0)))
        for r in recs_sorted:
            key = (
                r.get("para_idx"),
                r.get("cell_row_idx"),
                r.get("cell_col_idx"),
                r.get("cell_para_idx"),
            )
            if key in seen:
                # Continuation on a later page or another fragment of the
                # same paragraph — don't re-enumerate. We track only the
                # FIRST occurrence.
                continue
            seen[key] = len(enum)
            enum.append({
                "enum_i": len(enum),  # 0-based
                "page": int(pk),
                "y": r.get("y"),
                "x": r.get("x"),
                "text": r.get("text", ""),
                "text_y_off": r.get("text_y_off", 0.0),
                "key": key,
            })
    return enum


def _build_pair(k: int, wp: dict, ox: dict, prefix_ok: bool) -> dict:
    wt = normalize_text(wp.get("text", ""))
    ot = normalize_text(ox.get("text", ""))
    page_delta = ox["page"] - wp.get("page", -1)
    y_diff = None
    y_diff_visual = None
    if page_delta == 0 and ox["y"] is not None and wp.get("y") is not None:
        y_diff = round(ox["y"] - wp["y"], 2)
        y_diff_visual = round(ox["y"] + ox.get("text_y_off", 0.0) - wp["y"], 2)
    return {
        "k": k,
        "word_i": wp.get("i"),
        "word_page": wp.get("page"),
        "word_y": wp.get("y"),
        "word_text": wt[:30],
        "oxi_page": ox["page"],
        "oxi_y": ox["y"],
        "oxi_text": ot[:30],
        "oxi_key": list(ox["key"]),
        "prefix_ok": prefix_ok,
        "page_delta": page_delta,
        "y_diff": y_diff,
        "y_diff_visual": y_diff_visual,
    }


def text_prefix_similar(a: str, b: str, n_min: int = 4) -> bool:
    """True if the two normalized strings share their first n_min chars
    (or one is a prefix of the other when both are shorter than n_min)."""
    a = normalize_text(a)
    b = normalize_text(b)
    if not a and not b:
        return True
    if not a or not b:
        return False
    n = min(len(a), len(b), n_min)
    if n == 0:
        return True
    return a[:n] == b[:n]


def diff_doc(doc_id: str, word: dict, oxi: dict, verbose: bool = False,
             resync_window: int = 5) -> dict:
    word_paras_all = word.get("paragraphs", [])
    oxi_enum_all = enumerate_oxi(oxi)

    # Word's `Document.Paragraphs.Count` includes empty paragraphs (empty
    # cells, end-of-cell trailing paragraphs) that Oxi doesn't enumerate
    # as text-bearing tuples. Filter both to non-empty to align cardinalities.
    word_paras = [p for p in word_paras_all if normalize_text(p.get("text", ""))]
    oxi_enum = [e for e in oxi_enum_all if normalize_text(e.get("text", ""))]
    n_word_all = len(word_paras_all)
    n_oxi_all = len(oxi_enum_all)
    n_word_nonempty = len(word_paras)
    n_oxi_nonempty = len(oxi_enum)
    n_word = n_word_nonempty
    n_oxi = n_oxi_nonempty

    # Sliding-window aligner: walk both lists; when texts don't agree,
    # peek ahead up to `resync_window` on either side to find a match.
    # If Oxi has an extra paragraph, mark it as oxi_only and skip; vice
    # versa for Word. If no re-sync within the window, pair them
    # positionally and flag as `prefix_ok=False` (a true drift point).
    pairs = []
    word_only = []
    oxi_only = []
    i_w = 0
    i_o = 0
    while i_w < n_word and i_o < n_oxi:
        wp = word_paras[i_w]
        ox = oxi_enum[i_o]
        wt = normalize_text(wp.get("text", ""))
        ot = normalize_text(ox.get("text", ""))
        if text_prefix_similar(wt, ot):
            pairs.append(_build_pair(len(pairs), wp, ox, prefix_ok=True))
            i_w += 1
            i_o += 1
            continue
        # try forward resync — prefer skipping ONE on either side first.
        resync = None
        for d in range(1, resync_window + 1):
            if i_o + d < n_oxi:
                ot2 = normalize_text(oxi_enum[i_o + d].get("text", ""))
                if text_prefix_similar(wt, ot2):
                    resync = ("o", d)
                    break
            if i_w + d < n_word:
                wt2 = normalize_text(word_paras[i_w + d].get("text", ""))
                if text_prefix_similar(wt2, ot):
                    resync = ("w", d)
                    break
        if resync is None:
            # No re-sync — pair them anyway (drift point) and advance both.
            pairs.append(_build_pair(len(pairs), wp, ox, prefix_ok=False))
            i_w += 1
            i_o += 1
        elif resync[0] == "o":
            # Oxi has extra paragraphs; mark them as oxi-only.
            for d in range(resync[1]):
                e = oxi_enum[i_o + d]
                oxi_only.append({
                    "oxi_key": list(e["key"]),
                    "page": e["page"],
                    "y": e["y"],
                    "text": normalize_text(e.get("text", ""))[:30],
                })
            i_o += resync[1]
        else:
            # Word has extra paragraphs; mark them as word-only.
            for d in range(resync[1]):
                p = word_paras[i_w + d]
                word_only.append({
                    "word_i": p.get("i"),
                    "page": p.get("page"),
                    "y": p.get("y"),
                    "text": normalize_text(p.get("text", ""))[:30],
                })
            i_w += resync[1]
    # Tail
    while i_w < n_word:
        p = word_paras[i_w]
        word_only.append({
            "word_i": p.get("i"),
            "page": p.get("page"),
            "y": p.get("y"),
            "text": normalize_text(p.get("text", ""))[:30],
        })
        i_w += 1
    while i_o < n_oxi:
        e = oxi_enum[i_o]
        oxi_only.append({
            "oxi_key": list(e["key"]),
            "page": e["page"],
            "y": e["y"],
            "text": normalize_text(e.get("text", ""))[:30],
        })
        i_o += 1

    n_match = len(pairs)
    text_mismatches = [p for p in pairs if not p["prefix_ok"]]

    page_deltas = [p["page_delta"] for p in pairs]
    y_diffs = [p["y_diff"] for p in pairs if p["y_diff"] is not None]
    y_diffs_visual = [p["y_diff_visual"] for p in pairs if p["y_diff_visual"] is not None]
    n_zero = sum(1 for d in page_deltas if d == 0)
    n_prefix_ok = sum(1 for p in pairs if p["prefix_ok"])

    def stats(vals):
        if not vals:
            return {}
        s = sorted(vals)
        return {
            "n": len(s),
            "min": s[0],
            "p25": s[len(s)//4],
            "median": s[len(s)//2],
            "p75": s[3*len(s)//4],
            "max": s[-1],
            "mean": round(sum(s)/len(s), 3),
        }

    summary = {
        "doc_id": doc_id,
        "n_word_all": n_word_all,
        "n_oxi_all": n_oxi_all,
        "n_word_nonempty": n_word,
        "n_oxi_nonempty": n_oxi,
        "n_match": n_match,
        "n_word_only": len(word_only),
        "n_oxi_only": len(oxi_only),
        "n_prefix_ok": n_prefix_ok,
        "n_prefix_mismatch": n_match - n_prefix_ok,
        "n_page_zero": n_zero,
        "n_page_nonzero": n_match - n_zero,
        "page_delta_hist": dict(sorted(Counter(page_deltas).items())),
        "y_diff_raw_stats": stats(y_diffs),
        "y_diff_visual_stats": stats(y_diffs_visual),
        # Frac of pairs aligned both in page and prefix — confidence proxy.
        # Higher = matcher confident in this doc's drift numbers.
        "alignment_confidence": round(
            (sum(1 for p in pairs if p["page_delta"] == 0 and p["prefix_ok"]) / n_match)
            if n_match else 0.0, 4
        ),
    }

    detail = {
        **summary,
        "pairs": pairs,
        "word_only": word_only,
        "oxi_only": oxi_only,
    }

    if verbose:
        print(f"=== {doc_id} (positional) ===")
        print(f"  n_word_all={n_word_all} nonempty={n_word} | n_oxi_all={n_oxi_all} nonempty={n_oxi}")
        print(f"  matched={n_match}, word_only={len(word_only)}, oxi_only={len(oxi_only)}")
        print(f"  prefix_ok={n_prefix_ok}/{n_match}  ({summary['n_prefix_mismatch']} mismatches)")
        print(f"  page_zero={n_zero}/{n_match}  page_delta_hist={summary['page_delta_hist']}")
        print(f"  y_diff_raw    : {summary['y_diff_raw_stats']}")
        print(f"  y_diff_visual : {summary['y_diff_visual_stats']}")
        print(f"  alignment_confidence={summary['alignment_confidence']}")
        if text_mismatches[:10]:
            print(f"  First 10 prefix mismatches (alignment drift candidates):")
            for p in text_mismatches[:10]:
                print(f"    k={p['k']:>3} wi={p['word_i']:>3} pg(w={p['word_page']} o={p['oxi_page']}) "
                      f"w_text={p['word_text']!r}  o_text={p['oxi_text']!r}")

    return detail


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)

    args = sys.argv[1:]
    verbose = False
    if "--verbose" in args:
        verbose = True
        args.remove("--verbose")
    prefix = args[0] if args else None

    word_files = sorted(f for f in os.listdir(WORD_DIR) if f.endswith(".json") and not f.startswith("_"))
    doc_ids = [os.path.splitext(f)[0] for f in word_files]
    if prefix:
        doc_ids = [d for d in doc_ids if d.startswith(prefix)]
    if not doc_ids:
        print(f"no docs matched (prefix={prefix})", file=sys.stderr)
        return 2

    summary_rows = []
    skipped = []
    for doc_id in doc_ids:
        word = load_json(os.path.join(WORD_DIR, f"{doc_id}.json"))
        oxi = load_json(os.path.join(OXI_DIR, f"{doc_id}.json"))
        if word is None or oxi is None:
            skipped.append({"doc_id": doc_id, "has_word": word is not None, "has_oxi": oxi is not None})
            continue
        # Detect: oxi record schema must include cell identity (new in S177).
        first_recs = next(iter(oxi.get("pages", {}).values()), [])
        if first_recs and "cell_row_idx" not in first_recs[0]:
            skipped.append({"doc_id": doc_id, "reason": "oxi schema lacks cell_row_idx (re-run measure_pagination_oxi.py)"})
            continue
        detail = diff_doc(doc_id, word, oxi, verbose=verbose)
        out_path = os.path.join(OUT_DIR, f"{doc_id}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(detail, f, ensure_ascii=False, indent=2)
        row = {k: detail[k] for k in (
            "doc_id", "n_word_nonempty", "n_oxi_nonempty", "n_match",
            "n_word_only", "n_oxi_only",
            "n_prefix_ok", "n_prefix_mismatch", "n_page_zero", "n_page_nonzero",
            "alignment_confidence", "page_delta_hist",
        )}
        row["y_diff_visual_median"] = detail["y_diff_visual_stats"].get("median")
        row["y_diff_visual_mean"] = detail["y_diff_visual_stats"].get("mean")
        summary_rows.append(row)
        if not verbose:
            conf = detail["alignment_confidence"]
            print(f"  {doc_id}: matched={detail['n_match']}/{detail['n_word_nonempty']} "
                  f"prefix_ok={detail['n_prefix_ok']}/{detail['n_match']} "
                  f"page_zero={detail['n_page_zero']}/{detail['n_match']} "
                  f"conf={conf}")

    summary = {
        "n_total": len(summary_rows),
        "n_skipped": len(skipped),
        "docs": summary_rows,
        "skipped": skipped,
    }
    summary_path = os.path.join(OUT_DIR, "_summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"\nsummary -> {summary_path}")
    if summary_rows:
        confs = [r["alignment_confidence"] for r in summary_rows]
        mean_conf = sum(confs) / len(confs)
        print(f"  mean alignment_confidence: {mean_conf:.4f}")
        print(f"  docs with conf < 0.9: {sum(1 for c in confs if c < 0.9)}/{len(confs)}")
    if skipped:
        print(f"  skipped: {len(skipped)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
