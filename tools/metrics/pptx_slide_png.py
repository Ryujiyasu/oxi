# -*- coding: utf-8 -*-
"""Export a blind deck's slides to PNG with PowerPoint COM.

The corpus's truth is a PDF PowerPoint exported (`ppt_pdf/<doc>.pdf`), rastered
with pymupdf. That is one step removed from what PowerPoint puts on screen, and
at least one deck (31) writes **every text-showing operator with `ca 0`** -- an
invisible search layer -- while pymupdf's raster shows the text anyway. If MuPDF
is drawing what PowerPoint marks invisible, that deck's reference is a fiction.

`Slide.Export` asks PowerPoint itself to raster the slide, with no PDF in
between, so the two together settle it.

★★VERDICT (2026-08-29): the PDF is right and THIS is the unreliable one.
`Slide.Export` **does not load the presentation's embedded fonts**. On blind 31
slide 2 it sets a shape asking for the embedded "Open Sauce" in a CJK fallback
face -- the curly quotes come out as primes -- and wraps it into FOUR lines,
while the truth PDF, the editing window (screenshotted at 68% zoom) and Oxi all
give FIVE. The window is what a person sees, so the PDF matches PowerPoint and
this export does not. Scale is not the variable: 1440, 3000 and 6000px exports
all put the same 290.40pt of ink on line 1, and `WithWindow=True` changes
nothing.

So keep this for what it can still answer -- whether a raster difference is in
the PDF pipeline -- and never make it the truth for a deck that embeds fonts,
which in this corpus is every deck.

    python tools/metrics/pptx_slide_png.py 31            # every slide
    python tools/metrics/pptx_slide_png.py 31 --slides 2

Output: `ssim_pptx/ppt_png/<doc>/slide_sN.png`, sized to match the pymupdf
raster of the same page at 150 DPI so the two can be diffed pixel for pixel.

★NEVER run this while the renderer is producing PNGs (`pptx_com_render_must_not_overlap`).
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
PNG_DIR = ROOT / "ssim_pptx" / "ppt_png"
DPI = 150


def deck_path(doc: str) -> Path:
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    key = f"{int(doc):02d}"
    for item in manifest:
        if f"{item['idx']:02d}" == key:
            p = ROOT / "pptx" / item["local"]
            if not p.exists():
                sys.exit(f"deck file missing: {p}")
            return p
    sys.exit(f"no deck for {doc}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("doc")
    ap.add_argument("--slides", default="", help="comma list, 1-based; default all")
    ap.add_argument("--opens", type=int, default=1,
                    help="open the deck N times and export from the LAST open. "
                         "The first open of a session is cold and does not reach "
                         "the embedded faces (pptx_truth_pdf_first_open_is_cold)")
    ap.add_argument("--tag", default="", help="suffix for the output dir")
    args = ap.parse_args()

    key = f"{int(args.doc):02d}"
    src = deck_path(args.doc)
    out = PNG_DIR / (key + args.tag)
    out.mkdir(parents=True, exist_ok=True)
    want = [int(s) for s in args.slides.split(",") if s.strip()]

    import win32com.client
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for _ in range(args.opens - 1):
            warm = app.Presentations.Open(str(src.resolve()), WithWindow=False)
            warm.Close()
        pres = app.Presentations.Open(str(src.resolve()), WithWindow=False)
        try:
            w_pt = float(pres.PageSetup.SlideWidth)
            h_pt = float(pres.PageSetup.SlideHeight)
            w_px = int(round(w_pt * DPI / 72.0))
            h_px = int(round(h_pt * DPI / 72.0))
            n = pres.Slides.Count
            print(f"{key}: {n} slides, {w_pt}x{h_pt}pt -> {w_px}x{h_px}px", flush=True)
            for i in range(1, n + 1):
                if want and i not in want:
                    continue
                dst = out / f"slide_s{i}.png"
                pres.Slides(i).Export(str(dst), "PNG", w_px, h_px)
                print(f"  s{i} -> {dst.name}", flush=True)
        finally:
            pres.Close()
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
