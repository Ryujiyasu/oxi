"""Batch LLA comparison: Word vs Libra and Word vs Oxi side-by-side.

For each doc in --doc-ids (or all docs that already have a Word LLA JSON
cached), this script:
  - Reads (or computes) Word LLA JSON
  - Reads (or computes) Oxi  LLA JSON
  - Reads (or computes) Libra LLA JSON
  - Computes LCS-based line_text_match_rate for each pair
  - Reports the deltas per doc, plus a roll-up

Requires:
  - pipeline_data/libra_pdf/<doc_id>.pdf  (run render_libra.py first)
  - Word LLA JSONs cached or Word available for new ones
  - Oxi LLA JSONs cached or oxi-gdi-renderer + --dump-layout available

Cache layout (under pipeline_data/lla_compare/<doc_id>/):
  - libra_lla.json
  - word_lla.json   (reused from lla_canary_* if present, otherwise computed)
  - oxi_lla.json    (reused from lla_canary_* if present, otherwise computed)

Output:
  pipeline_data/lla_compare/_summary.json
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
from compute_lla import _lcs_len  # noqa: E402

REPO_ROOT = SCRIPT_DIR.parent.parent
PIPELINE_DATA = REPO_ROOT / "pipeline_data"
LIBRA_PDF_DIR = PIPELINE_DATA / "libra_pdf"
COMPARE_DIR = PIPELINE_DATA / "lla_compare"
SSIM_BASELINE = PIPELINE_DATA / "ssim_baseline.json"
LLA_CANARY_15 = PIPELINE_DATA / "lla_canary_15"
DOCX_DIR = REPO_ROOT / "tools" / "golden-test" / "documents" / "docx"
OXI_GDI = REPO_ROOT / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"

PY = sys.executable


def find_docx(doc_id: str) -> Path | None:
    candidates = list(DOCX_DIR.glob(f"{doc_id}*.docx"))
    if not candidates:
        return None
    # Prefer exact stem match if present
    for c in candidates:
        if c.stem == doc_id:
            return c
    return candidates[0]


def find_cached_word_lla(doc_id: str) -> Path | None:
    """Reuse lla_canary_*/<doc_id>__lla_word.json if present."""
    for parent in PIPELINE_DATA.glob("lla_canary_*"):
        cand = parent / f"{doc_id}__lla_word.json"
        if cand.is_file():
            return cand
    return None


def find_cached_oxi_lla(doc_id: str) -> Path | None:
    for parent in PIPELINE_DATA.glob("lla_canary_*"):
        cand = parent / f"{doc_id}__lla_oxi.json"
        if cand.is_file():
            return cand
    return None


def ensure_word_lla(doc_id: str, out_path: Path) -> bool:
    if out_path.is_file():
        return True
    cached = find_cached_word_lla(doc_id)
    if cached:
        shutil.copyfile(cached, out_path)
        return True
    docx = find_docx(doc_id)
    if not docx:
        return False
    pdf_cache = PIPELINE_DATA / "lla_pdf_cache" / f"{doc_id}.pdf"
    cmd = [PY, str(SCRIPT_DIR / "measure_lla_word.py"),
           str(docx), "-o", str(out_path), "--pdf-cache", str(pdf_cache)]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode == 0 and out_path.is_file()


def ensure_oxi_lla(doc_id: str, out_path: Path) -> bool:
    if out_path.is_file():
        return True
    cached = find_cached_oxi_lla(doc_id)
    if cached:
        shutil.copyfile(cached, out_path)
        return True
    if not OXI_GDI.is_file():
        return False
    docx = find_docx(doc_id)
    if not docx:
        return False
    layout = out_path.with_name(out_path.stem + "_layout.json")
    cmd1 = [str(OXI_GDI), str(docx), str(out_path.parent / doc_id),
            "150", f"--dump-layout={layout}"]
    proc = subprocess.run(cmd1, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if not layout.is_file():
        return False
    cmd2 = [PY, str(SCRIPT_DIR / "measure_lla_oxi.py"),
            str(layout), "-o", str(out_path)]
    proc = subprocess.run(cmd2, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode == 0 and out_path.is_file()


def ensure_libra_lla(doc_id: str, out_path: Path) -> bool:
    if out_path.is_file():
        return True
    pdf = LIBRA_PDF_DIR / f"{doc_id}.pdf"
    if not pdf.is_file():
        return False
    cmd = [PY, str(SCRIPT_DIR / "measure_lla_libra.py"),
           doc_id, "-o", str(out_path)]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode == 0 and out_path.is_file()


def lla_rate(word_pages: dict, target_pages: dict) -> tuple[int, int, float]:
    """Aggregate LCS rate across pages (same logic as compute_lla.diff_one)."""
    lcs = 0
    total = 0
    all_pages = sorted(set(word_pages) | set(target_pages), key=lambda k: int(k))
    for p in all_pages:
        wl = word_pages.get(p, [])
        tl = target_pages.get(p, [])
        lcs += _lcs_len(wl, tl)
        total += max(len(wl), len(tl))
    rate = lcs / total if total else 0.0
    return lcs, total, rate


def discover_doc_ids(prefix: str | None, baseline_only: bool, only_cached: bool, limit: int) -> list[str]:
    if only_cached:
        ids = set()
        for cand in PIPELINE_DATA.glob("lla_canary_*/*__lla_word.json"):
            ids.add(cand.name.replace("__lla_word.json", ""))
        ids = sorted(ids)
    elif baseline_only and SSIM_BASELINE.is_file():
        with SSIM_BASELINE.open(encoding="utf-8") as f:
            ids = sorted(json.load(f).keys())
    else:
        ids = sorted(p.stem for p in DOCX_DIR.glob("*.docx"))
    if prefix:
        ids = [d for d in ids if d.startswith(prefix)]
    if limit > 0:
        ids = ids[:limit]
    return ids


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--prefix", default=None)
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--baseline", action="store_true",
                    help="restrict to docs in ssim_baseline.json")
    ap.add_argument("--only-cached", action="store_true",
                    help="restrict to docs that already have a Word LLA JSON "
                         "cached under lla_canary_* (no new Word COM calls)")
    args = ap.parse_args()

    doc_ids = discover_doc_ids(args.prefix, args.baseline, args.only_cached, args.limit)
    if not doc_ids:
        sys.exit("no doc ids matched")

    print(f"# comparing LLA for {len(doc_ids)} doc(s)")
    COMPARE_DIR.mkdir(parents=True, exist_ok=True)

    records: list[dict] = []
    for i, doc_id in enumerate(doc_ids, start=1):
        rec: dict = {"doc_id": doc_id}
        per_doc = COMPARE_DIR / doc_id
        per_doc.mkdir(parents=True, exist_ok=True)

        word_p = per_doc / "word_lla.json"
        oxi_p = per_doc / "oxi_lla.json"
        libra_p = per_doc / "libra_lla.json"

        rec["word_ok"] = ensure_word_lla(doc_id, word_p)
        rec["oxi_ok"] = ensure_oxi_lla(doc_id, oxi_p)
        rec["libra_ok"] = ensure_libra_lla(doc_id, libra_p)

        if rec["word_ok"] and rec["oxi_ok"] and rec["libra_ok"]:
            word = json.loads(word_p.read_text(encoding="utf-8"))
            oxi = json.loads(oxi_p.read_text(encoding="utf-8"))
            libra = json.loads(libra_p.read_text(encoding="utf-8"))
            wp = word.get("pages", {})

            lcs_o, tot_o, rate_o = lla_rate(wp, oxi.get("pages", {}))
            lcs_l, tot_l, rate_l = lla_rate(wp, libra.get("pages", {}))
            rec.update({
                "word_pages": word.get("n_pages"),
                "oxi_pages": oxi.get("n_pages"),
                "libra_pages": libra.get("n_pages"),
                "oxi_lcs": lcs_o,
                "oxi_total": tot_o,
                "oxi_rate": round(rate_o, 4),
                "libra_lcs": lcs_l,
                "libra_total": tot_l,
                "libra_rate": round(rate_l, 4),
                "delta_libra_minus_oxi": round(rate_l - rate_o, 4),
            })
            verdict = ("Libra+" if rec["delta_libra_minus_oxi"] > 0.005
                       else "Oxi+" if rec["delta_libra_minus_oxi"] < -0.005
                       else "~")
            print(f"[{i:3}/{len(doc_ids)}] {doc_id:55.55s} "
                  f"oxi={rec['oxi_rate']*100:5.1f}%  libra={rec['libra_rate']*100:5.1f}%  "
                  f"delta={rec['delta_libra_minus_oxi']*100:+5.1f}%  {verdict}")
        else:
            missing = [k for k in ("word_ok", "oxi_ok", "libra_ok") if not rec[k]]
            print(f"[{i:3}/{len(doc_ids)}] {doc_id:55.55s} SKIP missing: {missing}")
        records.append(rec)

    ok = [r for r in records if "oxi_rate" in r]
    summary = {"n_total": len(records), "n_scored": len(ok)}
    if ok:
        summary.update({
            "mean_oxi_rate": round(sum(r["oxi_rate"] for r in ok) / len(ok), 4),
            "mean_libra_rate": round(sum(r["libra_rate"] for r in ok) / len(ok), 4),
            "mean_delta": round(sum(r["delta_libra_minus_oxi"] for r in ok) / len(ok), 4),
            "n_libra_better": sum(1 for r in ok if r["delta_libra_minus_oxi"] > 0.005),
            "n_oxi_better": sum(1 for r in ok if r["delta_libra_minus_oxi"] < -0.005),
            "n_tied": sum(1 for r in ok if abs(r["delta_libra_minus_oxi"]) <= 0.005),
        })
        print()
        print(f"# Mean LLA Oxi:   {summary['mean_oxi_rate']*100:.1f}%")
        print(f"# Mean LLA Libra: {summary['mean_libra_rate']*100:.1f}%")
        print(f"# Mean delta:     {summary['mean_delta']*100:+.1f}%  (positive = Libra closer to Word)")
        print(f"# Libra better: {summary['n_libra_better']}  Oxi better: {summary['n_oxi_better']}  tied: {summary['n_tied']}")

    out = COMPARE_DIR / "_summary.json"
    out.write_text(json.dumps({"summary": summary, "records": records}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n# wrote {out}")


if __name__ == "__main__":
    main()
