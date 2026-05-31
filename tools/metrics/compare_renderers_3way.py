"""Three-way renderer fidelity comparison against Word as ground truth.

For every (doc_id, page) in ssim_baseline.json, report SSIM(Word, X) for
X in {Oxi, LibreOffice, OnlyOffice}:
  - oxi_score: read from ssim_baseline.json (the live Phase-3 gate value)
  - libra_score: computed from pipeline_data/libra_png/...
  - oo_score:    computed from pipeline_data/onlyoffice_png/...

This does NOT modify ssim_baseline.json. It is a competitive sanity-check:
where do ALL third-party engines beat Oxi (= a real Oxi weakness worth
fixing), and where does Oxi already lead the field.

Pre-req:
  - tools/metrics/render_libra.py --baseline      (libra_png populated)
  - tools/metrics/render_onlyoffice.py --baseline (onlyoffice_png populated)

Output:
  pipeline_data/renderers_3way_ssim.json   (per-page records + summary)

Usage (from repo root):
    python tools/metrics/compare_renderers_3way.py
    python tools/metrics/compare_renderers_3way.py --prefix 0e7af
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
OO_PNG_DIR = PIPELINE_DATA / "onlyoffice_png"
SSIM_BASELINE = PIPELINE_DATA / "ssim_baseline.json"
OUT_PATH = PIPELINE_DATA / "renderers_3way_ssim.json"


def score(word_png: Path, other_png: Path) -> float:
    word = Image.open(word_png).convert("L")
    other = Image.open(other_png).convert("L")
    if other.size != word.size:
        other = other.resize(word.size, Image.LANCZOS)
    return float(ssim(np.array(word), np.array(other), data_range=255))


def mean(xs: list[float]) -> float:
    return sum(xs) / len(xs) if xs else float("nan")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--prefix", default=None)
    ap.add_argument("--out", default=str(OUT_PATH))
    args = ap.parse_args()

    if not SSIM_BASELINE.is_file():
        sys.exit(f"baseline missing: {SSIM_BASELINE}")
    baseline = json.loads(SSIM_BASELINE.read_text(encoding="utf-8"))

    doc_ids = sorted(baseline.keys())
    if args.prefix:
        doc_ids = [d for d in doc_ids if d.startswith(args.prefix)]

    records: list[dict] = []
    for doc_id in doc_ids:
        for page_key, oxi_score in baseline[doc_id].items():
            page_num = int(page_key)
            word_png = WORD_PNG_DIR / doc_id / f"page_{page_num:04d}.png"
            libra_png = LIBRA_PNG_DIR / doc_id / f"page_{page_num:04d}.png"
            oo_png = OO_PNG_DIR / doc_id / f"page_{page_num:04d}.png"
            rec = {"doc_id": doc_id, "page": page_num, "oxi_score": float(oxi_score)}
            if not word_png.is_file():
                rec["status"] = "word_png_missing"
                records.append(rec)
                continue
            rec["libra_score"] = score(word_png, libra_png) if libra_png.is_file() else None
            rec["oo_score"] = score(word_png, oo_png) if oo_png.is_file() else None
            rec["status"] = "ok"
            records.append(rec)

    full = [r for r in records if r["status"] == "ok"
            and r.get("libra_score") is not None and r.get("oo_score") is not None]

    summary = {
        "n_total_pages": len(records),
        "n_full_scored": len(full),
        "n_libra_missing": sum(1 for r in records if r["status"] == "ok" and r.get("libra_score") is None),
        "n_oo_missing": sum(1 for r in records if r["status"] == "ok" and r.get("oo_score") is None),
    }
    if full:
        summary["mean_oxi"] = round(mean([r["oxi_score"] for r in full]), 4)
        summary["mean_libra"] = round(mean([r["libra_score"] for r in full]), 4)
        summary["mean_oo"] = round(mean([r["oo_score"] for r in full]), 4)
        # Where every third-party engine beats Oxi (real Oxi weakness).
        both_beat = [r for r in full
                     if r["libra_score"] - r["oxi_score"] > 0.01
                     and r["oo_score"] - r["oxi_score"] > 0.01]
        # Where Oxi beats every third-party engine (Oxi lead).
        oxi_leads = [r for r in full
                     if r["oxi_score"] - r["libra_score"] > 0.01
                     and r["oxi_score"] - r["oo_score"] > 0.01]
        summary["n_both_beat_oxi"] = len(both_beat)
        summary["n_oxi_leads_field"] = len(oxi_leads)
        # rank: count pages where Oxi is best / worst of the three
        def best_of(r):
            return max([("oxi", r["oxi_score"]), ("libra", r["libra_score"]),
                        ("oo", r["oo_score"])], key=lambda t: t[1])[0]
        summary["n_oxi_best"] = sum(1 for r in full if best_of(r) == "oxi")
        summary["n_libra_best"] = sum(1 for r in full if best_of(r) == "libra")
        summary["n_oo_best"] = sum(1 for r in full if best_of(r) == "oo")

    out = {"summary": summary, "records": records}
    Path(args.out).write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"# scored {len(full)} pages fully (3 engines), {len(records)} total")
    if full:
        print(f"# mean SSIM vs Word:  Oxi={summary['mean_oxi']:.4f}  "
              f"Libra={summary['mean_libra']:.4f}  OnlyOffice={summary['mean_oo']:.4f}")
        print(f"# best-of-three page wins:  Oxi={summary['n_oxi_best']}  "
              f"Libra={summary['n_libra_best']}  OnlyOffice={summary['n_oo_best']}")
        print(f"# pages where BOTH third-party engines beat Oxi (real weakness): "
              f"{summary['n_both_beat_oxi']}")
        print(f"# pages where Oxi leads the whole field: {summary['n_oxi_leads_field']}")

        print("\n# Top 12 pages where BOTH Libra & OnlyOffice beat Oxi "
              "(min third-party advantage, sorted by how much):")
        bb = sorted(both_beat,
                    key=lambda r: -min(r["libra_score"] - r["oxi_score"],
                                       r["oo_score"] - r["oxi_score"]))[:12]
        for r in bb:
            print(f"  {r['doc_id']:46.46s} p.{r['page']:>3}  "
                  f"oxi={r['oxi_score']:.3f}  libra={r['libra_score']:.3f}  oo={r['oo_score']:.3f}")

        print("\n# Top 12 pages where Oxi leads the field (min lead):")
        ol = sorted(oxi_leads,
                    key=lambda r: -min(r["oxi_score"] - r["libra_score"],
                                       r["oxi_score"] - r["oo_score"]))[:12]
        for r in ol:
            print(f"  {r['doc_id']:46.46s} p.{r['page']:>3}  "
                  f"oxi={r['oxi_score']:.3f}  libra={r['libra_score']:.3f}  oo={r['oo_score']:.3f}")
    print(f"\n# wrote {args.out}")


if __name__ == "__main__":
    main()
