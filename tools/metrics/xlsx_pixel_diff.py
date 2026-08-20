"""Compare what Oxi draws for a worksheet against what Excel shows.

Excel's truth is its used range copied as a picture, which is the sheet as it
appears on screen. A printed PDF was tried first and abandoned: it obeys print
areas, page breaks and shrink-to-fit, none of which Oxi models, so a workbook
whose print area covered one row was compared against a whole sheet, and any
sheet Excel broke across pages could not be compared at all.

Oxi's side comes from oxi-xlsx-renderer. Both are cropped to the ink they hold
before being compared, since the two disagree about the margin around a sheet.

Usage:
    python tools/metrics/xlsx_pixel_diff.py <dir-or-file.xlsx>
        [--dpi 96] [--out pipeline_data/xlsx_diff] [--reuse] [--limit N]
        [--baseline path] [--update-baseline]

Reports per-file SSIM, a mean, and what moved against the baseline, and writes
a three-panel PNG per file: Excel on the left, Oxi in the middle, the difference
on the right.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity

REPO = Path(__file__).resolve().parents[2]
RENDERER = (
    REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
)
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"


def excel_shots(pairs: list[tuple[Path, Path]], out: Path, chunk: int = 25) -> set[Path]:
    """Copies each sheet's used range as a picture, keeping Excel open throughout.

    Starting Excel once rather than per file is the difference between minutes
    and an hour over a few hundred workbooks. The batch is split into chunks so
    a workbook Excel will not let go of costs one chunk rather than the run, and
    anything already taken is left alone so an interrupted run can pick up where
    it stopped.
    """
    pending = [(source, dest) for source, dest in pairs if not dest.exists()]
    already = {dest for _, dest in pairs if dest.exists()}
    if not pending:
        return already
    out.mkdir(parents=True, exist_ok=True)

    for first in range(0, len(pending), chunk):
        batch = pending[first : first + chunk]
        listing = out / "_batch.txt"
        # PowerShell runs from its own directory, so a relative path would not
        # find the workbook.
        listing.write_text(
            chr(10).join(
                str(source.resolve()) + chr(9) + str(destination.resolve())
                for source, destination in batch
            ),
            encoding="utf-8",
        )
        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-File",
                    str(SHOOTER),
                    "-ListFile",
                    str(listing.resolve()),
                ],
                capture_output=True,
                text=True,
                timeout=60 * len(batch),
            )
        except subprocess.TimeoutExpired:
            # Excel is stuck on one of these. Kill it and carry on: what it did
            # take is on disk, and a later run will retry the rest.
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    "Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force",
                ],
                capture_output=True,
            )
            print(f"  Excel stalled on the chunk starting {batch[0][0].name}")
        listing.unlink(missing_ok=True)
        done = sum(1 for _, dest in pending[: first + len(batch)] if dest.exists())
        print(f"  took {done}/{len(pending)}")

    return already | {dest for _, dest in pending if dest.exists()}


def crop_to_ink(image: Image.Image, pad: int = 2) -> Image.Image:
    """Trims the blank margin, so two pictures of a sheet line up."""
    grey = np.asarray(image.convert("L"))
    dark = np.argwhere(grey < 200)
    if dark.size == 0:
        return image
    top, left = dark.min(axis=0)
    bottom, right = dark.max(axis=0)
    return image.crop(
        (
            max(int(left) - pad, 0),
            max(int(top) - pad, 0),
            min(int(right) + pad + 1, image.width),
            min(int(bottom) + pad + 1, image.height),
        )
    )


def compare(truth: Image.Image, ours: Image.Image, panel: Path) -> float:
    truth, ours = crop_to_ink(truth), crop_to_ink(ours)
    width = max(truth.width, ours.width)
    height = max(truth.height, ours.height)

    def onto_canvas(image: Image.Image) -> Image.Image:
        canvas = Image.new("RGB", (width, height), (255, 255, 255))
        canvas.paste(image, (0, 0))
        return canvas

    truth, ours = onto_canvas(truth), onto_canvas(ours)
    left = np.asarray(truth.convert("L"))
    right = np.asarray(ours.convert("L"))
    # The same SSIM the docx gate reports, so the two numbers mean the same
    # thing. Unlike that gate this one pads rather than resizes, because a
    # sheet drawn at the wrong size differs for a reason worth seeing.
    score = float(structural_similarity(left, right, data_range=255))

    difference = Image.fromarray(
        255 - np.abs(left.astype(int) - right.astype(int)).astype(np.uint8)
    )
    strip = Image.new("RGB", (width * 3 + 20, height), (240, 240, 240))
    strip.paste(truth, (0, 0))
    strip.paste(ours, (width + 10, 0))
    strip.paste(difference.convert("RGB"), (width * 2 + 20, 0))
    panel.parent.mkdir(parents=True, exist_ok=True)
    strip.save(panel)
    return score


def report_against_baseline(scores: dict[str, float], baseline: Path) -> None:
    if not baseline.exists():
        print(f"no baseline at {baseline}")
        return
    held = json.loads(baseline.read_text())["scores"]
    shared = [stem for stem in scores if stem in held]
    if not shared:
        print("the baseline holds none of these workbooks")
        return
    moved = [
        (stem, scores[stem] - held[stem])
        for stem in shared
        if abs(scores[stem] - held[stem]) > 1e-4
    ]
    better = [pair for pair in moved if pair[1] > 0]
    worse = [pair for pair in moved if pair[1] < 0]
    was = sum(held[stem] for stem in shared) / len(shared)
    now = sum(scores[stem] for stem in shared) / len(shared)
    print(
        f"against the baseline: {len(better)} improved, {len(worse)} regressed,"
        f" mean {was:.4f} -> {now:.4f} ({now - was:+.4f}) over {len(shared)}"
    )
    for stem, delta in sorted(moved, key=lambda pair: pair[1])[:5]:
        print(f"   {delta:+.4f}  {stem}")
    for stem, delta in sorted(moved, key=lambda pair: -pair[1])[:5]:
        print(f"   {delta:+.4f}  {stem}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("target", type=Path)
    parser.add_argument("--dpi", type=float, default=96.0)
    parser.add_argument("--out", type=Path, default=REPO / "pipeline_data" / "xlsx_diff")
    parser.add_argument(
        "--reuse",
        action="store_true",
        help="compare against the pictures already on disk, without opening Excel",
    )
    parser.add_argument("--limit", type=int, help="stop after this many workbooks")
    parser.add_argument(
        "--baseline",
        type=Path,
        default=REPO / "pipeline_data" / "xlsx_ssim_baseline.json",
        help="scores to compare this run against",
    )
    parser.add_argument(
        "--update-baseline",
        action="store_true",
        help="write this run's scores as the new baseline",
    )
    args = parser.parse_args()

    if not RENDERER.exists():
        print(f"build the renderer first: cargo build --release in {RENDERER.parents[2]}")
        return 1

    sources = sorted(args.target.glob("*.xlsx")) if args.target.is_dir() else [args.target]
    sources = [path for path in sources if not path.name.startswith("~$")]
    if args.limit:
        sources = sources[: args.limit]
    if not sources:
        print("no workbooks to compare")
        return 1

    pairs = [(source, args.out / f"{source.stem}.excel.png") for source in sources]
    if args.reuse:
        taken = {destination for _, destination in pairs if destination.exists()}
        print(f"reusing {len(taken)} of {len(sources)} picture(s)")
    else:
        taken = excel_shots(pairs, args.out)
        print(f"Excel gave up {len(taken)} of {len(sources)} picture(s)")

    scores: dict[str, float] = {}
    for source in sources:
        stem = source.stem
        truth_png = args.out / f"{stem}.excel.png"
        ours_png = args.out / f"{stem}.oxi.png"
        if truth_png not in taken:
            continue
        drawn = subprocess.run(
            [str(RENDERER), str(source), str(ours_png), str(args.dpi)],
            capture_output=True,
            text=True,
        )
        if not ours_png.exists():
            print(f"  Oxi could not draw {source.name}: {drawn.stderr.strip()[:120]}")
            continue
        score = compare(
            Image.open(truth_png).convert("RGB"),
            Image.open(ours_png),
            args.out / f"{stem}.compare.png",
        )
        scores[stem] = score
        print(f"  {stem:40s} SSIM {score:.4f}")

    if not scores:
        print("nothing compared")
        return 1
    mean = sum(scores.values()) / len(scores)
    summary = {"mean": mean, "scores": scores, "dpi": args.dpi}
    (args.out / "_summary.json").write_text(json.dumps(summary, indent=2))
    print(f"\nmean SSIM {mean:.4f} over {len(scores)} workbook(s)")

    report_against_baseline(scores, args.baseline)
    if args.update_baseline:
        args.baseline.parent.mkdir(parents=True, exist_ok=True)
        args.baseline.write_text(json.dumps(summary, indent=2))
        print(f"baseline written to {args.baseline}")
    print(f"panels and summary in {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
