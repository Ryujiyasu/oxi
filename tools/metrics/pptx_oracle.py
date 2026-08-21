# -*- coding: utf-8 -*-
"""Reference-renderer oracle over the pptx dev corpus (the docx bug-finder,
brought to slides at the user's direction 2026-08-21).

An oracle render of each deck is scored against PowerPoint's own PDF with
EXACTLY the harness conventions of `pptx_ssim_floor.py` (150 DPI, RGB SSIM,
LANCZOS resize, slide-weighted mean), so its numbers sit next to Oxi's on the
same scale. A slide where the oracle lands close to PowerPoint and Oxi does
not is a fixable Oxi bug with a working reference to diff against.

Oracles:
  libra    LibreOffice Impress: soffice --headless --convert-to pdf, then
           rasterize (the docx render_libra.py recipe)
  silurus  @silurus/ooxml pptx viewer via tools/browser-oracle (harness_pptx),
           driven by pptx_browser_oracle.py

Usage:
    python tools/metrics/pptx_oracle.py libra [--decks d09,d20] [--rerender]
    python tools/metrics/pptx_oracle.py report [--vs ss3]

`report` reads every oracle's stored JSON plus the Oxi ssim_floor result named
by --vs and prints the per-slide "oracle beats Oxi" queue.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from skimage.metrics import structural_similarity

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev"
PPTX_DIR = DEV / "pptx"
PDF_DIR = DEV / "pdf"
REF_ROOT = DEV / "ppt_png"
ORACLE_ROOT = DEV / "oracle"
SOFFICE = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
DPI = 150


def decks(selector: str | None) -> list[Path]:
    all_decks = sorted(PPTX_DIR.glob("*.pptx"))
    if not selector:
        return all_decks
    wanted = {s.strip() for s in selector.split(",") if s.strip()}
    return [p for p in all_decks if p.stem.split("__")[0] in wanted or p.stem in wanted]


def reference_pages(pdf_path: Path) -> list[Path]:
    """The ppt_png cache pptx_ssim_floor.py keeps; build it if missing."""
    cache = REF_ROOT / pdf_path.stem
    pngs = sorted(cache.glob("p*.png"), key=lambda p: int(p.stem[1:]))
    if pngs:
        return pngs
    cache.mkdir(parents=True, exist_ok=True)
    pdf = pymupdf.open(pdf_path)
    try:
        for index in range(pdf.page_count):
            pix = pdf[index].get_pixmap(matrix=pymupdf.Matrix(DPI / 72, DPI / 72), alpha=False)
            pix.save(cache / f"p{index + 1}.png")
    finally:
        pdf.close()
    return sorted(cache.glob("p*.png"), key=lambda p: int(p.stem[1:]))


def score(reference: np.ndarray, candidate: np.ndarray) -> float:
    if candidate.shape != reference.shape:
        candidate = np.asarray(
            Image.fromarray(candidate).resize(
                (reference.shape[1], reference.shape[0]), Image.Resampling.LANCZOS
            )
        )
    return float(structural_similarity(reference, candidate, channel_axis=2, data_range=255))


def libra_render(pptx: Path, rerender: bool) -> list[Path] | None:
    """Impress render of one deck -> oracle/libra_png/<deck>/p<N>.png."""
    pdf_cache = ORACLE_ROOT / "libra_pdf" / f"{pptx.stem}.pdf"
    png_dir = ORACLE_ROOT / "libra_png" / pptx.stem
    if rerender:
        pdf_cache.unlink(missing_ok=True)
        shutil.rmtree(png_dir, ignore_errors=True)
    pngs = sorted(png_dir.glob("p*.png"), key=lambda p: int(p.stem[1:]))
    if pngs:
        return pngs
    if not pdf_cache.is_file():
        pdf_cache.parent.mkdir(parents=True, exist_ok=True)
        # Fresh profile so a hung previous run's lock never wedges this one.
        with tempfile.TemporaryDirectory(prefix="soffice_profile_") as profile, \
                tempfile.TemporaryDirectory(prefix="soffice_out_") as outdir:
            cmd = [
                str(SOFFICE), "--headless", "--norestore", "--nologo",
                "--nofirststartwizard",
                f"-env:UserInstallation=file:///{Path(profile).as_posix()}",
                "--convert-to", "pdf", "--outdir", outdir, str(pptx),
            ]
            try:
                subprocess.run(cmd, capture_output=True, timeout=180)
            except subprocess.TimeoutExpired:
                return None
            produced = list(Path(outdir).glob("*.pdf"))
            if not produced:
                return None
            shutil.copy(produced[0], pdf_cache)
    png_dir.mkdir(parents=True, exist_ok=True)
    pdf = pymupdf.open(pdf_cache)
    try:
        for index in range(pdf.page_count):
            pix = pdf[index].get_pixmap(matrix=pymupdf.Matrix(DPI / 72, DPI / 72), alpha=False)
            pix.save(png_dir / f"p{index + 1}.png")
    finally:
        pdf.close()
    return sorted(png_dir.glob("p*.png"), key=lambda p: int(p.stem[1:]))


def measure(oracle: str, pptx: Path, rerender: bool) -> dict | None:
    pdf_path = PDF_DIR / f"{pptx.stem}.pdf"
    if not pdf_path.is_file():
        return None
    if oracle == "libra":
        pages = libra_render(pptx, rerender)
    else:
        # silurus pages are produced by pptx_browser_oracle.py into the same
        # oracle/<name>_png/<deck>/p<N>.png layout; here they are only scored.
        png_dir = ORACLE_ROOT / f"{oracle}_png" / pptx.stem
        pages = sorted(png_dir.glob("p*.png"), key=lambda p: int(p.stem[1:]))
    if not pages:
        return None
    refs = reference_pages(pdf_path)
    scores = []
    for index in range(len(refs)):
        if index >= len(pages):
            break
        reference = np.asarray(Image.open(refs[index]).convert("RGB"))
        candidate = np.asarray(Image.open(pages[index]).convert("RGB"))
        scores.append(round(score(reference, candidate), 6))
    if not scores:
        return None
    return {
        "deck": pptx.stem,
        "ppt_pages": len(refs),
        "oracle_pages": len(pages),
        "common_mean": round(sum(scores) / len(scores), 6),
        "slides": scores,
    }


def run(oracle: str, selector: str | None, rerender: bool) -> None:
    rows = []
    targets = decks(selector)
    for pptx in targets:
        row = measure(oracle, pptx, rerender)
        if row is None:
            print(f"  FAIL {pptx.stem[:50]}")
            continue
        rows.append(row)
        print(
            f"  {row['deck'][:40]:40s} {row['common_mean']:.4f}  "
            f"pages {row['oracle_pages']}/{row['ppt_pages']}"
        )
    total = sum(len(r["slides"]) for r in rows)
    mean = sum(s for r in rows for s in r["slides"]) / total if total else None
    print(f"\n{oracle}: {len(rows)}/{len(targets)} decks, {total} slides, mean {mean:.6f}"
          if mean is not None else f"\n{oracle}: nothing scored")
    ORACLE_ROOT.mkdir(parents=True, exist_ok=True)
    out = ORACLE_ROOT / f"{oracle}.json"
    out.write_text(
        json.dumps({"oracle": oracle, "mean": mean, "rows": rows}, indent=1),
        encoding="utf-8",
    )
    print(f"wrote {out}")


def report(vs: str) -> None:
    oxi = json.loads(
        (DEV / "ssim_floor" / f"{vs}.json").read_text(encoding="utf-8")
    )
    oxi_rows = {r["deck"]: r for r in oxi["rows"]}
    for path in sorted(ORACLE_ROOT.glob("*.json")):
        data = json.loads(path.read_text(encoding="utf-8"))
        name = data["oracle"]
        gaps = []
        for row in data["rows"]:
            o = oxi_rows.get(row["deck"])
            if not o:
                continue
            for i, s in enumerate(row["slides"]):
                if i < len(o["slides"]):
                    gaps.append((s - o["slides"][i], row["deck"], i + 1, s, o["slides"][i]))
        gaps.sort(reverse=True)
        wins = sum(1 for g in gaps if g[0] > 0.005)
        print(f"\n== {name} vs {vs}: {wins}/{len(gaps)} slides where {name} beats Oxi by >0.005")
        for delta, deck, page, s_o, s_x in gaps[:25]:
            if delta <= 0.005:
                break
            print(f"  +{delta:.4f}  {deck[:40]:40s} s{page:<3d} {name}={s_o:.4f} oxi={s_x:.4f}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("oracle", choices=["libra", "silurus", "report"])
    ap.add_argument("--decks", default=None)
    ap.add_argument("--rerender", action="store_true")
    ap.add_argument("--vs", default="ss3", help="ssim_floor result to compare against")
    args = ap.parse_args()
    if args.oracle == "report":
        report(args.vs)
        return
    if args.oracle == "libra" and not SOFFICE.is_file():
        sys.exit(f"soffice not found: {SOFFICE}")
    run(args.oracle, args.decks, args.rerender)


if __name__ == "__main__":
    main()
