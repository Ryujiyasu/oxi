"""Compute SSIM(Word, Libra) per page across the SSIM baseline, in the
same way pipeline/ssim_calculator.py computes SSIM(Word, Oxi).

The output is a side-by-side comparison: for every (doc_id, page) in
ssim_baseline.json we report:
  - oxi_score:   the SSIM(Word, Oxi) score (read from ssim_baseline.json)
  - libra_score: the SSIM(Word, Libra) score (computed here)
  - delta:       libra_score - oxi_score  (positive => Libra is closer to Word)

We do NOT modify ssim_baseline.json; this is a comparison report.

Output:
  pipeline_data/libra_vs_oxi_ssim.json   (per-page records + summary)

Pre-req: tools/metrics/render_libra.py --baseline must have been run so
that pipeline_data/libra_png/<doc_id>/page_NNNN.png exist.

Usage (from repo root):
    python tools/metrics/compute_ssim_libra.py
    python tools/metrics/compute_ssim_libra.py --prefix 04b88e
    python tools/metrics/compute_ssim_libra.py --limit 10
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

REPO_ROOT = Path(__file__).resolve().parents[2]
PIPELINE_DATA = REPO_ROOT / "pipeline_data"
WORD_PNG_DIR = PIPELINE_DATA / "word_png"
LIBRA_PNG_DIR = PIPELINE_DATA / "libra_png"
SSIM_BASELINE = PIPELINE_DATA / "ssim_baseline.json"
OUT_PATH = PIPELINE_DATA / "libra_vs_oxi_ssim.json"


def load_gray(path: Path, target_size: tuple[int, int] | None = None) -> np.ndarray:
    img = Image.open(path).convert("L")
    if target_size and img.size != target_size:
        img = img.resize(target_size, Image.LANCZOS)
    return np.array(img)


def compute_one_score(word_png: Path, libra_png: Path) -> float:
    word = Image.open(word_png).convert("L")
    libra = Image.open(libra_png).convert("L")
    # Normalize to the same shape; pipeline/ssim_calculator.py uses target=word size
    if libra.size != word.size:
        libra = libra.resize(word.size, Image.LANCZOS)
    return float(ssim(np.array(word), np.array(libra), data_range=255))


def normalize_page_key(k: str) -> int:
    return int(k)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--prefix", default=None)
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--out", default=str(OUT_PATH))
    args = ap.parse_args()

    if not SSIM_BASELINE.is_file():
        sys.exit(f"baseline missing: {SSIM_BASELINE}")
    with SSIM_BASELINE.open(encoding="utf-8") as f:
        baseline = json.load(f)

    doc_ids = sorted(baseline.keys())
    if args.prefix:
        doc_ids = [d for d in doc_ids if d.startswith(args.prefix)]
    if args.limit > 0:
        doc_ids = doc_ids[: args.limit]

    print(f"# comparing {len(doc_ids)} docs ({sum(len(baseline[d]) for d in doc_ids)} pages)")

    records: list[dict] = []
    n_libra_missing = 0
    n_word_missing = 0
    n_scored = 0
    for doc_id in doc_ids:
        pages = baseline[doc_id]
        # baseline page keys are sometimes "1"/"2", sometimes "0001"/"0002"
        for page_key, oxi_score in pages.items():
            page_num = normalize_page_key(page_key)
            word_png = WORD_PNG_DIR / doc_id / f"page_{page_num:04d}.png"
            libra_png = LIBRA_PNG_DIR / doc_id / f"page_{page_num:04d}.png"
            rec = {
                "doc_id": doc_id,
                "page": page_num,
                "oxi_score": float(oxi_score),
            }
            if not word_png.is_file():
                rec["status"] = "word_png_missing"
                n_word_missing += 1
                records.append(rec)
                continue
            if not libra_png.is_file():
                rec["status"] = "libra_png_missing"
                n_libra_missing += 1
                records.append(rec)
                continue
            try:
                libra_score = compute_one_score(word_png, libra_png)
            except Exception as e:
                rec["status"] = "error"
                rec["error"] = str(e)
                records.append(rec)
                continue
            rec["libra_score"] = libra_score
            rec["delta"] = libra_score - rec["oxi_score"]
            rec["status"] = "ok"
            records.append(rec)
            n_scored += 1

    # summary
    ok = [r for r in records if r["status"] == "ok"]
    summary = {
        "n_total_pages": len(records),
        "n_scored": n_scored,
        "n_libra_missing": n_libra_missing,
        "n_word_missing": n_word_missing,
    }
    if ok:
        summary.update({
            "mean_oxi_score": round(sum(r["oxi_score"] for r in ok) / len(ok), 4),
            "mean_libra_score": round(sum(r["libra_score"] for r in ok) / len(ok), 4),
            "mean_delta_libra_minus_oxi": round(sum(r["delta"] for r in ok) / len(ok), 4),
            "n_libra_better": sum(1 for r in ok if r["delta"] > 0.001),
            "n_oxi_better": sum(1 for r in ok if r["delta"] < -0.001),
            "n_within_001": sum(1 for r in ok if abs(r["delta"]) <= 0.001),
        })

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as f:
        json.dump({"summary": summary, "records": records}, f, ensure_ascii=False, indent=2)

    print(f"\n# wrote {out_path}")
    print(f"# scored: {n_scored}, libra missing: {n_libra_missing}, word missing: {n_word_missing}")
    if ok:
        print(f"# mean Oxi   SSIM: {summary['mean_oxi_score']:.4f}")
        print(f"# mean Libra SSIM: {summary['mean_libra_score']:.4f}")
        sign = "+" if summary["mean_delta_libra_minus_oxi"] >= 0 else ""
        print(f"# mean delta:      {sign}{summary['mean_delta_libra_minus_oxi']:.4f}  (positive = Libra closer to Word)")
        print(f"# Libra better: {summary['n_libra_better']}  Oxi better: {summary['n_oxi_better']}  ~tied: {summary['n_within_001']}")

        # Top-10 most-different pages either way
        worst_for_libra = sorted([r for r in ok if r["delta"] < 0], key=lambda r: r["delta"])[:10]
        worst_for_oxi = sorted([r for r in ok if r["delta"] > 0], key=lambda r: -r["delta"])[:10]
        if worst_for_libra:
            print("\n# Pages where Oxi beats Libra (top 10):")
            for r in worst_for_libra:
                print(f"  {r['doc_id']:50.50s} p.{r['page']:>3}  "
                      f"oxi={r['oxi_score']:.4f}  libra={r['libra_score']:.4f}  delta={r['delta']:+.4f}")
        if worst_for_oxi:
            print("\n# Pages where Libra beats Oxi (top 10):")
            for r in worst_for_oxi:
                print(f"  {r['doc_id']:50.50s} p.{r['page']:>3}  "
                      f"oxi={r['oxi_score']:.4f}  libra={r['libra_score']:.4f}  delta={r['delta']:+.4f}")


if __name__ == "__main__":
    main()
