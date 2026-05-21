"""Phase A: Oxi vs Word ground truth diff.

For each paragraph in word_layout_truth, find the corresponding Oxi paragraph
(by text-prefix match), then compare per-line y positions.

Reports:
- Per-paragraph dy (start_y diff)
- Per-paragraph line count mismatch
- Per-line dy (line top diff)
- Disagreement summary (counts by magnitude / category)

Usage:
    python tools/metrics/word_truth_diff.py <doc_prefix>     # single doc detail
    python tools/metrics/word_truth_diff.py --all            # all docs with truth + oxi summary
"""
import json, os, sys
from pathlib import Path
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
TRUTH_DIR = REPO_ROOT / "pipeline_data" / "word_layout_truth"
OXI_DIR = REPO_ROOT / "pipeline_data" / "pagination_oxi"


def load_truth(doc_id):
    p = TRUTH_DIR / f"{doc_id}.json"
    if not p.exists():
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def load_oxi(doc_id):
    p = OXI_DIR / f"{doc_id}.json"
    if not p.exists():
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def normalize_text(s):
    return s.replace("　", " ").replace("\r", "").replace("\x07", "").strip()


def match_oxi_para(truth_para, oxi_records_by_page):
    """Find Oxi paragraph matching truth paragraph by text-prefix on same page (±1 page)."""
    target_text = normalize_text(truth_para["text_prefix"])
    if not target_text:
        return None
    target_page = truth_para["start_page"]
    if target_page is None:
        return None

    for delta in [0, -1, 1, -2, 2]:
        page = target_page + delta
        page_str = str(page)
        if page_str not in oxi_records_by_page:
            continue
        for rec in oxi_records_by_page[page_str]:
            rec_text = normalize_text(rec.get("text", ""))
            if not rec_text:
                continue
            # Match if either prefix of the other
            min_len = min(len(rec_text), len(target_text), 12)
            if rec_text[:min_len] == target_text[:min_len]:
                return (page, rec)
    return None


def diff_doc(doc_id, verbose=True):
    truth = load_truth(doc_id)
    if truth is None:
        return None
    oxi = load_oxi(doc_id)
    if oxi is None:
        return None

    # Build oxi_records_by_page (oxi has page numbers as string keys under 'pages')
    oxi_by_page = oxi.get("pages", oxi) if isinstance(oxi, dict) else {}

    rows = []
    for tpara in truth["paragraphs"]:
        if not tpara["text_prefix"]:
            continue
        match = match_oxi_para(tpara, oxi_by_page)
        if match is None:
            rows.append({
                "i": tpara["i"],
                "truth_page": tpara["start_page"],
                "truth_y": tpara["start_y"],
                "truth_n_lines": tpara["n_lines"],
                "oxi_page": None,
                "oxi_y": None,
                "matched": False,
                "text_prefix": tpara["text_prefix"][:50],
            })
            continue
        oxi_page, oxi_rec = match
        rows.append({
            "i": tpara["i"],
            "truth_page": tpara["start_page"],
            "truth_y": tpara["start_y"],
            "truth_n_lines": tpara["n_lines"],
            "oxi_page": oxi_page,
            "oxi_y": oxi_rec.get("y"),
            "oxi_text_y_off": oxi_rec.get("text_y_off", 0.0),
            "matched": True,
            "in_table": tpara["in_table"],
            "fn_ref_count": tpara["fn_ref_count"],
            "text_prefix": tpara["text_prefix"][:50],
        })

    # Summary stats
    matched = [r for r in rows if r["matched"]]
    n_matched = len(matched)
    n_unmatched = len(rows) - n_matched

    page_diffs = []
    y_diffs = []
    y_diffs_visual = []  # using text_y_off correction
    for r in matched:
        page_diffs.append(r["oxi_page"] - r["truth_page"])
        dy = r["oxi_y"] - r["truth_y"]
        y_diffs.append(dy)
        # Visual y (Oxi line top + text_y_off → text actual top)
        dy_visual = (r["oxi_y"] + r["oxi_text_y_off"]) - r["truth_y"]
        y_diffs_visual.append(dy_visual)

    def stats(vals):
        if not vals:
            return {}
        s = sorted(vals)
        return {
            "n": len(vals),
            "min": s[0],
            "p25": s[len(s)//4],
            "median": s[len(s)//2],
            "p75": s[3*len(s)//4],
            "max": s[-1],
            "mean": sum(vals)/len(vals),
        }

    n_page_mismatch = sum(1 for p in page_diffs if p != 0)
    n_large_y = sum(1 for dy in y_diffs if abs(dy) >= 5.0)
    n_large_y_visual = sum(1 for dy in y_diffs_visual if abs(dy) >= 5.0)

    summary = {
        "doc_id": doc_id,
        "n_paragraphs_truth": len(truth["paragraphs"]),
        "n_matched": n_matched,
        "n_unmatched": n_unmatched,
        "n_page_mismatch": n_page_mismatch,
        "n_large_y_diff_raw": n_large_y,
        "n_large_y_diff_visual": n_large_y_visual,
        "y_diff_raw": stats(y_diffs),
        "y_diff_visual": stats(y_diffs_visual),
        "page_diff_hist": dict((d, page_diffs.count(d)) for d in set(page_diffs)),
    }

    if verbose:
        print(f"=== {doc_id} ===")
        print(f"truth paras: {len(truth['paragraphs'])}, matched: {n_matched}, unmatched: {n_unmatched}")
        print(f"page mismatch: {n_page_mismatch}")
        print(f"y_diff_raw    : {summary['y_diff_raw']}")
        print(f"y_diff_visual : {summary['y_diff_visual']}")
        print(f"large_y_diff: raw={n_large_y} visual={n_large_y_visual}")
        print(f"page_diff_hist: {summary['page_diff_hist']}")
        # Show paragraphs with significant disagreement
        print()
        print(f"Significant disagreements:")
        print(f"  {'i':>4} {'pg':>4} {'tr_y':>7} {'ox_y':>7} {'dy_raw':>7} {'tyoff':>6} {'dy_vis':>7} {'page_d':>7} text")
        sig = []
        for r in matched:
            dy_raw = r["oxi_y"] - r["truth_y"]
            dy_vis = (r["oxi_y"] + r["oxi_text_y_off"]) - r["truth_y"]
            pd = r["oxi_page"] - r["truth_page"]
            if abs(dy_raw) >= 5.0 or abs(dy_vis) >= 5.0 or pd != 0:
                sig.append((r, dy_raw, dy_vis, pd))
        for r, dyr, dyv, pd in sig[:20]:
            print(f"  {r['i']:>4} {r['truth_page']:>4} {r['truth_y']:>7.2f} {r['oxi_y']:>7.2f} {dyr:+7.2f} {r['oxi_text_y_off']:>6.2f} {dyv:+7.2f} {pd:+7} {r['text_prefix'][:40]!r}")

    return summary


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return

    if args[0] == "--all":
        truth_files = sorted(TRUTH_DIR.glob("*.json"))
        print(f"{'doc_id':<14} {'n_para':>6} {'matched':>7} {'pg_mis':>6} {'med_raw':>7} {'med_vis':>7} {'big_raw':>7} {'big_vis':>7}")
        for f in truth_files:
            doc_id = f.stem
            s = diff_doc(doc_id, verbose=False)
            if s is None:
                continue
            med_raw = s["y_diff_raw"].get("median", 0)
            med_vis = s["y_diff_visual"].get("median", 0)
            print(f"{doc_id:<14} {s['n_paragraphs_truth']:>6} {s['n_matched']:>7} {s['n_page_mismatch']:>6} {med_raw:>+7.2f} {med_vis:>+7.2f} {s['n_large_y_diff_raw']:>7} {s['n_large_y_diff_visual']:>7}")
    else:
        for prefix in args:
            truth_files = list(TRUTH_DIR.glob(f"{prefix}*.json"))
            for f in truth_files:
                diff_doc(f.stem)


if __name__ == "__main__":
    main()
