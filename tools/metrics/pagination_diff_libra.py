"""Run the Phase 1 pagination diff against Libra (instead of Oxi).

Reuses the diff_doc() logic from pagination_diff.py — the matching is
text-prefix based, so Libra's per-page line records can be substituted
for Oxi's per-page paragraph records without code change.

Reads:
  pipeline_data/pagination_word/<doc_id>.json    (existing, Word ground truth)
  pipeline_data/pagination_libra/<doc_id>.json   (from measure_pagination_libra.py)

Writes:
  pipeline_data/pagination_diff_libra/<doc_id>.json   (per-doc detail)
  pipeline_data/pagination_diff_libra/_summary.json   (cross-doc)

ALSO produces a side-by-side report by reading the Oxi-side summary
(pipeline_data/pagination_diff/_summary.json) and joining on doc_id.

Run from repo root:
    python tools/metrics/pagination_diff_libra.py            # all docs
    python tools/metrics/pagination_diff_libra.py 04b88e     # prefix
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
import pagination_diff  # noqa: E402

REPO_ROOT = SCRIPT_DIR.parent.parent
PIPELINE_DATA = REPO_ROOT / "pipeline_data"
WORD_DIR = PIPELINE_DATA / "pagination_word"
LIBRA_DIR = PIPELINE_DATA / "pagination_libra"
OUT_DIR = PIPELINE_DATA / "pagination_diff_libra"
OXI_SUMMARY = PIPELINE_DATA / "pagination_diff" / "_summary.json"


def main() -> int:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    prefix = sys.argv[1] if len(sys.argv) > 1 else None
    word_files = sorted(f for f in WORD_DIR.iterdir() if f.suffix == ".json" and not f.name.startswith("_"))
    doc_ids = [p.stem for p in word_files]
    if prefix:
        doc_ids = [d for d in doc_ids if d.startswith(prefix)]
    if not doc_ids:
        print(f"no docs matched (WORD_DIR={WORD_DIR}, prefix={prefix})", file=sys.stderr)
        return 2

    # Load Oxi summary for side-by-side join
    oxi_by_doc: dict[str, dict] = {}
    if OXI_SUMMARY.is_file():
        oxi_sum = json.loads(OXI_SUMMARY.read_text(encoding="utf-8"))
        for d in oxi_sum.get("docs", []):
            oxi_by_doc[d["doc_id"]] = d

    # Word pagination JSON is keyed by 12-char hash (e.g. "04b88e7e0b25")
    # Libra pagination JSON is keyed by full docx stem (e.g. "04b88e7e0b25_index-19")
    # Build a lookup: hash_prefix -> full Libra json path
    libra_by_prefix: dict[str, Path] = {}
    for p in LIBRA_DIR.glob("*.json"):
        prefix = p.stem.split("_")[0]
        libra_by_prefix[prefix] = p

    summary = []
    n_pass = 0
    n_fail = 0
    n_skipped = 0

    for doc_id in doc_ids:
        word_path = WORD_DIR / f"{doc_id}.json"
        libra_path = libra_by_prefix.get(doc_id)
        if libra_path is None or not libra_path.is_file():
            n_skipped += 1
            continue
        word = json.loads(word_path.read_text(encoding="utf-8"))
        libra = json.loads(libra_path.read_text(encoding="utf-8"))

        result = pagination_diff.diff_doc(doc_id, word, libra)
        # don't write the per-match detail to disk — big and not needed for summary
        per_doc = {k: v for k, v in result.items() if k != "matches"}
        per_doc["sample_unmatched"] = [m for m in result["matches"] if not m["matched"]][:5]
        (OUT_DIR / f"{doc_id}.json").write_text(
            json.dumps(per_doc, ensure_ascii=False, indent=2), encoding="utf-8")

        if result["pass"]:
            n_pass += 1
        else:
            n_fail += 1

        # join with Oxi
        oxi = oxi_by_doc.get(doc_id, {})
        summary.append({
            "doc_id": doc_id,
            "libra_pass": result["pass"],
            "libra_score": result["score"],
            "libra_n_matched": result["n_matched"],
            "libra_page_count_delta": result["page_count_delta"],
            "oxi_pass": oxi.get("pass"),
            "oxi_score": oxi.get("score"),
            "oxi_page_count_delta": oxi.get("page_count_delta"),
        })

    scored = [s for s in summary if s["oxi_score"] is not None]
    n_total = len(summary)
    mean_libra = (sum(s["libra_score"] for s in summary) / n_total) if n_total else 0.0
    if scored:
        mean_oxi = sum(s["oxi_score"] for s in scored) / len(scored)
        delta = mean_libra - mean_oxi
        libra_pass = sum(1 for s in scored if s["libra_pass"])
        oxi_pass = sum(1 for s in scored if s["oxi_pass"])
    else:
        mean_oxi = None
        delta = None
        libra_pass = sum(1 for s in summary if s["libra_pass"])
        oxi_pass = None

    out_summary = {
        "n_total": n_total,
        "n_libra_pass": n_pass,
        "n_libra_fail": n_fail,
        "n_skipped": n_skipped,
        "libra_pass_rate": round(n_pass / n_total, 4) if n_total else 0.0,
        "mean_libra_score": round(mean_libra, 4),
        "n_oxi_pass_in_join": oxi_pass,
        "mean_oxi_score_in_join": round(mean_oxi, 4) if mean_oxi is not None else None,
        "mean_delta_libra_minus_oxi": round(delta, 4) if delta is not None else None,
        "docs": summary,
    }
    (OUT_DIR / "_summary.json").write_text(
        json.dumps(out_summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"# Libra pagination: {n_pass}/{n_total} pass ({out_summary['libra_pass_rate']*100:.1f}%), "
          f"mean score {out_summary['mean_libra_score']:.4f}")
    if mean_oxi is not None:
        print(f"# Oxi pagination (joined {len(scored)} docs): {oxi_pass}/{len(scored)} pass, "
              f"mean score {mean_oxi:.4f}")
        sign = "+" if delta >= 0 else ""
        print(f"# delta libra - oxi: {sign}{delta:.4f} (positive = Libra closer to Word)")
    print(f"\n# wrote {OUT_DIR / '_summary.json'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
