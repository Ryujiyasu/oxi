# -*- coding: utf-8 -*-
"""PPTX dev-corpus SSIM floor: which decks / slides are worst RIGHT NOW.

The docx side has `ssim_floor.py`; the pptx side had nothing -- every past
session measured the dev corpus with a throw-away scratch script, so there was
no way to ask "what is the floor at HEAD" without rewriting the harness. This
is that tool, with the same conventions as the blind-50 harness
(`pipeline_data/pptx_benchmark/_measure_ssim_pptx.py`):

  reference = the PowerPoint PDF of the same deck (pipeline_data/pptx_benchmark/
              dev/pdf/<deck>.pdf), rasterised at 150 DPI
  candidate = oxi-pptx-renderer PNGs at 150 DPI
  score     = skimage SSIM on RGB (channel_axis=2, data_range=255), candidate
              LANCZOS-resized to the reference when the raster sizes differ

Reported per deck: the mean over slides PowerPoint and Oxi both produced
(`common_mean`), the same sum divided by max(ppt_pages, oxi_pages)
(`penalized_mean`, so a deck that loses slides cannot hide it), and the worst
single slide. The corpus number is slide-weighted, which is what every dev
ship-note in the ledger quotes.

Renders are cached per --tag, so an A/B is two runs with different tags and
different --env, and neither has to re-render the other's arm:

    python tools/metrics/pptx_ssim_floor.py --tag head
    python tools/metrics/pptx_ssim_floor.py --tag off --env OXI_FOO_DISABLE=1
    python tools/metrics/pptx_ssim_floor.py --ab off head

Usage:
    python tools/metrics/pptx_ssim_floor.py [--tag NAME] [--decks d01,d04]
                                            [--env K=V ...] [--rerender]
                                            [--jobs N] [--json OUT.json]
                                            [--ab BEFORE AFTER]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import zipfile
import subprocess
import sys
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image
from skimage.metrics import structural_similarity

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).resolve().parents[2]
DEV = REPO_ROOT / "pipeline_data" / "pptx_benchmark" / "dev"
PPTX_DIR = DEV / "pptx"
PDF_DIR = DEV / "pdf"
PNG_ROOT = DEV / "oxi_png"
RESULT_ROOT = DEV / "ssim_floor"
REF_ROOT = DEV / "ppt_png"
# MuPDF keeps shared state, so the reference rasters are produced under
# one lock and then reused from disk.
LOCK = threading.Lock()
OXI_EXE = REPO_ROOT / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"
DPI = 150


def decks(selector: str | None) -> list[Path]:
    all_decks = sorted(PPTX_DIR.glob("*.pptx"))
    if not selector:
        return all_decks
    wanted = {s.strip() for s in selector.split(",") if s.strip()}
    return [p for p in all_decks if p.stem.split("__")[0] in wanted or p.stem in wanted]


def deck_id(pptx: Path) -> str:
    return pptx.stem


def deck_slide_count(pptx: Path) -> int:
    """How many slides the deck declares, so a partial cache can be spotted."""
    try:
        with zipfile.ZipFile(pptx) as z:
            names = [n for n in z.namelist()
                     if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
        return len(names)
    except Exception:
        return 0


def render(pptx: Path, tag: str, env_pairs: dict[str, str], rerender: bool) -> dict:
    """Render one deck to PNG_ROOT/<tag>/<deck>/slide_sN.png (cached)."""
    out_dir = PNG_ROOT / tag / deck_id(pptx)
    if rerender and out_dir.exists():
        shutil.rmtree(out_dir)
    existing = list(out_dir.glob("slide_s*.png"))
    if existing:
        # ★A partial cache is worse than no cache: it reads as a complete deck
        # and silently scores the slides that happened to finish. An
        # interrupted render is the usual way to get one.
        want = deck_slide_count(pptx)
        if want and len(existing) < want:
            print(f"  {deck_id(pptx)}: cache has {len(existing)}/{want} slides "
                  f"-- discarding and re-rendering")
            shutil.rmtree(out_dir)
        else:
            return {"deck": deck_id(pptx), "ok": True, "cached": True}
    out_dir.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    env.update(env_pairs)
    try:
        proc = subprocess.run(
            [str(OXI_EXE), str(pptx), str(out_dir / "slide"), str(DPI)],
            capture_output=True,
            text=True,
            errors="replace",
            timeout=1800,
            env=env,
        )
    except subprocess.TimeoutExpired:
        return {"deck": deck_id(pptx), "ok": False, "error": "timeout"}
    produced = list(out_dir.glob("slide_s*.png"))
    if proc.returncode != 0 or not produced:
        detail = (proc.stderr or proc.stdout or "")[-200:]
        return {"deck": deck_id(pptx), "ok": False, "error": f"rc={proc.returncode} {detail}"}
    return {"deck": deck_id(pptx), "ok": True, "cached": False}


def reference_pages(pdf_path: Path) -> list[Path]:
    """Rasterise the PowerPoint PDF once and cache the pages as PNG.

    Rendering the reference inside the scoring loop made the harness
    NON-DETERMINISTIC: two runs over BYTE-IDENTICAL Oxi PNGs disagreed by
    3e-5 on d02, because the pages were rasterised concurrently and MuPDF's
    state is shared. Caching removes the noise (an A/B is now exactly
    reproducible) and makes every later run much cheaper, since the truth
    side never changes.
    """
    cache = REF_ROOT / pdf_path.stem
    pngs = sorted(cache.glob("p*.png"), key=lambda p: int(p.stem[1:]))
    if pngs:
        return pngs
    cache.mkdir(parents=True, exist_ok=True)
    with LOCK:
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


def measure(pptx: Path, tag: str) -> dict | None:
    """Score one deck against its PowerPoint PDF. None = no ground truth."""
    did = deck_id(pptx)
    pdf_path = PDF_DIR / f"{did}.pdf"
    if not pdf_path.is_file():
        return None
    png_dir = PNG_ROOT / tag / did
    pngs = sorted(png_dir.glob("slide_s*.png"), key=lambda p: int(p.stem.split("_s")[1]))
    refs = reference_pages(pdf_path)
    scores: list[float] = []
    for index in range(len(refs)):
        if index >= len(pngs):
            break
        reference = np.asarray(Image.open(refs[index]).convert("RGB"))
        candidate = np.asarray(Image.open(pngs[index]).convert("RGB"))
        scores.append(score(reference, candidate))
    ppt_pages = len(refs)
    denominator = max(ppt_pages, len(pngs))
    worst = min(range(len(scores)), key=lambda i: scores[i]) if scores else None
    return {
        "deck": did,
        "ppt_pages": ppt_pages,
        "oxi_pages": len(pngs),
        "page_delta": len(pngs) - ppt_pages,
        "common_pages": len(scores),
        "common_mean": round(sum(scores) / len(scores), 6) if scores else None,
        "penalized_mean": round(sum(scores) / denominator, 6) if denominator else None,
        "worst_slide": (worst + 1) if worst is not None else None,
        "worst_score": round(scores[worst], 6) if worst is not None else None,
        "slides": [round(s, 6) for s in scores],
    }


def unchanged_vs(tag: str, other: str, deck: str) -> bool:
    """True when this deck's renders are byte-identical between two tags."""
    a = PNG_ROOT / tag / deck
    b = PNG_ROOT / other / deck
    if not a.is_dir() or not b.is_dir():
        return False
    an = sorted(p.name for p in a.glob("slide_s*.png"))
    if an != sorted(p.name for p in b.glob("slide_s*.png")):
        return False
    return all(
        hashlib.sha256((a / n).read_bytes()).digest()
        == hashlib.sha256((b / n).read_bytes()).digest()
        for n in an
    )


def run(args) -> dict:
    targets = decks(args.decks)
    if not targets:
        sys.exit("no decks matched")
    env_pairs = dict(pair.split("=", 1) for pair in args.env)
    print(f"decks={len(targets)} tag={args.tag} env={env_pairs or '-'}", flush=True)
    cached_already = sum(
        1
        for t in targets
        if list((PNG_ROOT / args.tag / deck_id(t)).glob("slide_s*.png"))
    )
    if cached_already and not args.rerender:
        # A tag that already has renders is REUSED silently, and a tag name
        # recycled from an earlier session then reports that session's build.
        # It cost a false -0.066 on d24 (2026-08-20) before the numbers were
        # traced back to renders nobody had made today.
        print(
            f"  ! {cached_already}/{len(targets)} decks REUSED from an existing "
            f"'{args.tag}' render -- pass --rerender or pick a fresh tag",
            flush=True,
        )
    # An A/B only needs the decks whose pixels moved: byte-identical renders
    # score identically by construction, and SSIM over 886 slides is the whole
    # cost of a run. Rows for the unchanged decks are taken from the reference
    # result so the corpus number stays complete.
    reuse: dict[str, dict] = {}
    if args.vs:
        prior = json.loads((RESULT_ROOT / f"{args.vs}.json").read_text(encoding="utf-8"))
        reuse = {r["deck"]: r for r in prior["rows"]}

    with ThreadPoolExecutor(max_workers=args.jobs) as pool:
        futures = {
            pool.submit(render, p, args.tag, env_pairs, args.rerender): p for p in targets
        }
        for future in as_completed(futures):
            info = future.result()
            if not info["ok"]:
                print(f"  RENDER FAIL {info['deck']}: {info.get('error')}", flush=True)

    rows: list[dict] = []
    to_score = targets
    if reuse:
        keep = [p for p in targets if deck_id(p) in reuse and unchanged_vs(args.tag, args.vs, deck_id(p))]
        rows.extend(reuse[deck_id(p)] for p in keep)
        to_score = [p for p in targets if p not in set(keep)]
        print(
            f"  reusing {len(keep)} byte-identical deck scores from '{args.vs}', "
            f"scoring {len(to_score)}",
            flush=True,
        )
    with ThreadPoolExecutor(max_workers=args.jobs) as pool:
        # skimage's SSIM releases the GIL inside numpy, so scoring threads do
        # overlap; the sequential version took longer to score 886 slides than
        # to render them.
        futures = {pool.submit(measure, p, args.tag): p for p in to_score}
        for future in as_completed(futures):
            row = future.result()
            if row is None:
                print(f"  SKIP {deck_id(futures[future])}: no PowerPoint PDF", flush=True)
                continue
            rows.append(row)
            print(
                f"  {row['deck'][:34]:34s} {row['common_mean']}  "
                f"pages {row['oxi_pages']}/{row['ppt_pages']}  "
                f"worst s{row['worst_slide']}={row['worst_score']}",
                flush=True,
            )
    rows.sort(key=lambda r: r["deck"])

    total_slides = sum(r["common_pages"] for r in rows)
    corpus = (
        sum(r["common_mean"] * r["common_pages"] for r in rows if r["common_mean"] is not None)
        / total_slides
        if total_slides
        else None
    )
    penalized_total = sum(max(r["ppt_pages"], r["oxi_pages"]) for r in rows)
    corpus_pen = (
        sum(
            r["penalized_mean"] * max(r["ppt_pages"], r["oxi_pages"])
            for r in rows
            if r["penalized_mean"] is not None
        )
        / penalized_total
        if penalized_total
        else None
    )
    summary = {
        "tag": args.tag,
        "env": env_pairs,
        "decks": len(rows),
        "slides": total_slides,
        "corpus_mean": round(corpus, 6) if corpus is not None else None,
        "corpus_penalized_mean": round(corpus_pen, 6) if corpus_pen is not None else None,
        "page_delta_decks": [r["deck"] for r in rows if r["page_delta"] != 0],
        "rows": rows,
    }

    print(f"\nCORPUS n={len(rows)} decks / {total_slides} slides")
    print(f"  slide-weighted mean   {summary['corpus_mean']}")
    print(f"  penalized mean        {summary['corpus_penalized_mean']}")
    if summary["page_delta_decks"]:
        print(f"  slide-count mismatch  {', '.join(summary['page_delta_decks'])}")
    print("\nFLOOR (worst decks first)")
    for row in sorted(rows, key=lambda r: r["common_mean"] if r["common_mean"] is not None else 2):
        print(
            f"  {row['common_mean']}  {row['deck'][:44]:44s} "
            f"{row['common_pages']:3d} slides  worst s{row['worst_slide']}={row['worst_score']}"
        )

    RESULT_ROOT.mkdir(parents=True, exist_ok=True)
    out = Path(args.json) if args.json else RESULT_ROOT / f"{args.tag}.json"
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\nwrote {out}")
    return summary


def ab(before_tag: str, after_tag: str) -> None:
    """Compare two stored runs, per deck and per slide."""
    a = json.loads((RESULT_ROOT / f"{before_tag}.json").read_text(encoding="utf-8"))
    b = json.loads((RESULT_ROOT / f"{after_tag}.json").read_text(encoding="utf-8"))
    a_rows = {r["deck"]: r for r in a["rows"]}
    b_rows = {r["deck"]: r for r in b["rows"]}
    shared = [d for d in a_rows if d in b_rows]
    improved = regressed = 0
    print(f"A/B  before={before_tag}  after={after_tag}\n")
    for d in sorted(shared, key=lambda d: b_rows[d]["common_mean"] - a_rows[d]["common_mean"]):
        delta = b_rows[d]["common_mean"] - a_rows[d]["common_mean"]
        if abs(delta) < 5e-5:
            continue
        improved += delta > 0
        regressed += delta < 0
        print(
            f"  {delta:+.4f}  {d[:44]:44s} "
            f"{a_rows[d]['common_mean']:.4f} -> {b_rows[d]['common_mean']:.4f}  "
            f"({b_rows[d]['common_pages']} slides)"
        )
    slides = sum(b_rows[d]["common_pages"] for d in shared)
    weighted_a = sum(a_rows[d]["common_mean"] * a_rows[d]["common_pages"] for d in shared)
    weighted_b = sum(b_rows[d]["common_mean"] * b_rows[d]["common_pages"] for d in shared)
    print(
        f"\nTOTAL {slides} slides  {weighted_a/slides:.6f} -> {weighted_b/slides:.6f} "
        f"({(weighted_b-weighted_a)/slides:+.6f})   {improved} improved / {regressed} regressed"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tag", default="head", help="cache + result name for this arm")
    parser.add_argument("--decks", default=None, help="comma-separated deck ids, e.g. d04,d10")
    parser.add_argument("--env", action="append", default=[], help="K=V passed to the renderer")
    parser.add_argument("--rerender", action="store_true", help="drop the PNG cache for this tag")
    parser.add_argument("--jobs", type=int, default=4)
    parser.add_argument("--json", default=None)
    parser.add_argument("--ab", nargs=2, metavar=("BEFORE", "AFTER"), default=None)
    parser.add_argument(
        "--vs",
        default=None,
        help="reuse this tag's scores for decks whose renders are byte-identical",
    )
    args = parser.parse_args()

    if args.ab:
        ab(*args.ab)
        return
    if not OXI_EXE.is_file():
        sys.exit(f"renderer not built: {OXI_EXE}")
    run(args)


if __name__ == "__main__":
    main()
