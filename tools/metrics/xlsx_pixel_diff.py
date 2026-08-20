"""Compare what Oxi draws for a worksheet against what Excel prints.

Excel's truth is a PDF from ExportAsFixedFormat, rasterised here. Oxi's side
comes from oxi-xlsx-renderer. The two are cropped to the ink they hold before
being compared, because Excel prints onto a full page while Oxi draws only the
used range.

Usage:
    python tools/metrics/xlsx_pixel_diff.py <dir-or-file.xlsx> [--dpi 96]
                                            [--out pipeline_data/xlsx_diff]

Reports per-file SSIM and a mean, and writes a three-panel PNG per file:
Excel on the left, Oxi in the middle, the difference on the right.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

import numpy as np

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover - reported at runtime
    fitz = None
from PIL import Image
from skimage.metrics import structural_similarity

REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
EXPORTER = Path(__file__).resolve().parent / "_xlsx_export_pdf.ps1"


def excel_pdfs(
    pairs: list[tuple[Path, Path]], out: Path, chunk: int = 25
) -> set[Path]:
    """Prints a batch of workbooks to PDF, keeping Excel open throughout.

    Starting Excel once rather than per file is the difference between minutes
    and an hour over a few hundred workbooks. The batch is split into chunks so
    a workbook Excel will not let go of costs one chunk rather than the run, and
    anything already printed is left alone so an interrupted run can pick up
    where it stopped.
    """
    pending = [(source, dest) for source, dest in pairs if not dest.exists()]
    already = {dest for _, dest in pairs if dest.exists()}
    if not pending:
        return already
    out.mkdir(parents=True, exist_ok=True)

    for start in range(0, len(pending), chunk):
        batch = pending[start : start + chunk]
        listing = out / "_batch.txt"
        # PowerShell runs from its own directory, so a relative path would not
        # find the workbook.
        body = [
            str(source.resolve()) + chr(9) + str(destination.resolve())
            for source, destination in batch
        ]
        listing.write_text(chr(10).join(body), encoding="utf-8")
        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-File",
                    str(EXPORTER),
                    "-ListFile",
                    str(listing.resolve()),
                ],
                capture_output=True,
                text=True,
                timeout=60 * len(batch),
            )
        except subprocess.TimeoutExpired:
            # Excel is stuck on one of these. Kill it and carry on: the files it
            # did print are on disk, and a later run will retry the rest.
            subprocess.run(
                ["powershell", "-NoProfile", "-Command",
                 "Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force"],
                capture_output=True,
            )
            print(f"  Excel stalled on the chunk starting {batch[0][0].name}")
        listing.unlink(missing_ok=True)
        done = sum(1 for _, dest in pending[: start + len(batch)] if dest.exists())
        print(f"  printed {done}/{len(pending)}")

    return already | {dest for _, dest in pending if dest.exists()}


def rasterise(pdf: Path, dpi: float) -> Image.Image:
    if fitz is None:
        raise SystemExit("PyMuPDF is needed to read the PDF Excel prints")
    with fitz.open(pdf) as document:
        page = document[0]
        pixmap = page.get_pixmap(dpi=int(dpi))
        return Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)


def crop_to_ink(image: Image.Image, pad: int = 2) -> Image.Image:
    """Trims the white margin, so a printed page and a drawn range line up."""
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
    # printed page and a drawn range differ in size for a reason worth seeing.
    score = float(structural_similarity(left, right, data_range=255))

    difference = Image.fromarray(255 - np.abs(left.astype(int) - right.astype(int)).astype(np.uint8))
    strip = Image.new("RGB", (width * 3 + 20, height), (240, 240, 240))
    strip.paste(truth, (0, 0))
    strip.paste(ours, (width + 10, 0))
    strip.paste(difference.convert("RGB"), (width * 2 + 20, 0))
    panel.parent.mkdir(parents=True, exist_ok=True)
    strip.save(panel)
    return score


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("target", type=Path)
    parser.add_argument("--dpi", type=float, default=96.0)
    parser.add_argument("--out", type=Path, default=REPO / "pipeline_data" / "xlsx_diff")
    args = parser.parse_args()

    if not RENDERER.exists():
        print(f"build the renderer first: cargo build --release in {RENDERER.parents[2]}")
        return 1

    sources = (
        sorted(args.target.glob("*.xlsx")) if args.target.is_dir() else [args.target]
    )
    sources = [path for path in sources if not path.name.startswith("~$")]
    if not sources:
        print("no workbooks to compare")
        return 1

    printed = excel_pdfs(
        [(source, args.out / f"{source.stem}.pdf") for source in sources], args.out
    )
    print(f"Excel printed {len(printed)} of {len(sources)} workbook(s)")

    scores: dict[str, float] = {}
    for source in sources:
        stem = source.stem
        pdf = args.out / f"{stem}.pdf"
        ours_png = args.out / f"{stem}.oxi.png"
        if pdf not in printed:
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
            rasterise(pdf, args.dpi),
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
    print(f"panels and summary in {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
