# -*- coding: utf-8 -*-
"""Where PowerPoint puts each CHARACTER, asked of PowerPoint.

`TextRange.Characters(i, 1).BoundLeft` answers per glyph, with no PDF and no
raster in the way. That matters more than it sounds, because the truth PDF is
not a per-glyph oracle: PowerPoint restates a line there through a `Tf` size
that is not the declared one, a per-run `Tc`, and sparse integer `TJ`
(`read_pptx_drawgrid_com.py`).

What this printed the first time it was pointed at things (2026-09-02):

    Arial 12pt, 'n' x 40      step 6.62504, slope over 40 glyphs 6.62500
                              (the master unit exactly; the design advance is
                              6.67383 and is not what PowerPoint uses)

    d35 s25, Open Sans 12pt   12 distinct steps, 0 of them off the 1/8pt grid
                              'o' 6.750  'u' 6.875  ' ' 3.125 -- which is what
                              the engine computes for the same characters

    the same line vs its PDF  worst divergence 0.230pt

★That last number closed a queue. `pptx_editor_glyph_sweep.py` charged exactly
0.230pt to the engine on that line -- so its "advance offenders" are measuring
the PDF writer disagreeing with PowerPoint, not the engine disagreeing with
anything. Roughly 0.2-0.25pt is that instrument's FLOOR, and pushing its
offender count toward zero would be optimising against the export.

With `--pdf` this prints that comparison, so the floor can be re-checked rather
than remembered.

★Must not overlap the renderer (`pptx_render_not_parallel_safe`) or another COM
session (`pptx_com_render_must_not_overlap`).

    python tools/metrics/pptx_char_pos_com.py d35 25 "You don" --pdf
    python tools/metrics/pptx_char_pos_com.py 21 12 "Gray"
"""
from __future__ import annotations

import argparse
import glob
import sys
from pathlib import Path

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
BENCH = REPO / "pipeline_data" / "pptx_benchmark"

# One master unit, the unit PowerPoint measures and draws in.
PER_PT = 8.0


def find_deck(name: str) -> tuple[Path, Path | None]:
    """A dev deck by its `dNN` name, or a blind deck by its index."""
    if name.lower().startswith("d"):
        pptx = sorted(glob.glob(str(BENCH / "dev" / "pptx" / f"{name}__*.pptx")))
        pdf = sorted(glob.glob(str(BENCH / "dev" / "pdf" / f"{name}__*.pdf")))
    else:
        pptx = sorted(glob.glob(str(BENCH / "pptx" / f"{int(name):02d}__*.pptx")))
        pdf = sorted(glob.glob(str(BENCH / "ssim_pptx" / "ppt_pdf" / f"{int(name):02d}.pdf")))
    if not pptx:
        sys.exit(f"no deck matching {name!r}")
    return Path(pptx[0]), (Path(pdf[0]) if pdf else None)


def pdf_line(pdf: Path, page_no: int, text: str) -> list[float] | None:
    """The same characters' pen positions out of the truth PDF, or None."""
    import pymupdf

    doc = pymupdf.open(pdf)
    chars = []
    for b in doc[page_no - 1].get_text("rawdict")["blocks"]:
        for line in b.get("lines", []):
            for span in line.get("spans", []):
                for ch in span.get("chars", []):
                    chars.append((ch["c"], ch["origin"][0]))
    doc.close()
    probe = text[: min(len(text), 40)]
    for i in range(len(chars) - len(probe)):
        if all(chars[i + k][0] == probe[k] for k in range(len(probe))):
            return [chars[i + k][1] for k in range(min(len(text), len(chars) - i))]
    return None


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("deck", help="a dev deck like d35, or a blind index like 21")
    ap.add_argument("slide", type=int)
    ap.add_argument("needle", help="text that identifies the shape")
    ap.add_argument("--pdf", action="store_true",
                    help="also compare against the truth PDF, to re-check its floor")
    ap.add_argument("--chars", type=int, default=0, help="print the first N steps")
    args = ap.parse_args()

    src, pdf = find_deck(args.deck)
    print(f"{src.name[:56]}  slide {args.slide}")
    app = win32com.client.Dispatch("PowerPoint.Application")
    found = None
    try:
        pres = app.Presentations.Open(str(src.resolve()), WithWindow=False)
        try:
            sl = pres.Slides(args.slide)
            for i in range(1, sl.Shapes.Count + 1):
                sh = sl.Shapes(i)
                try:
                    if not sh.HasTextFrame or not sh.TextFrame.HasText:
                        continue
                    tr = sh.TextFrame.TextRange
                except Exception:
                    continue
                if args.needle not in tr.Text:
                    continue
                txt = tr.Text
                end = txt.find("\r")
                end = len(txt) if end < 0 else end
                found = (
                    txt[:end],
                    [tr.Characters(k, 1).BoundLeft for k in range(1, end + 1)],
                    tr.Characters(1, 1).Font.Name,
                    tr.Characters(1, 1).Font.Size,
                )
                break
        finally:
            pres.Saved = True
            pres.Close()
    finally:
        app.Quit()

    if not found:
        sys.exit(f"no shape on slide {args.slide} contains {args.needle!r}")
    text, xs, face, size = found
    # ★A paragraph that WRAPS steps backwards at every line start: reading the
    # 196-character version of the line above put -570.3751 in the step list
    # and called it "off grid". Only steps that move forward within one line
    # are advances.
    raw = [(k, xs[k + 1] - xs[k]) for k in range(len(xs) - 1)]
    steps = [round(v, 4) for _, v in raw if 0.0 < v < size * 3.0]
    wraps = len(raw) - len(steps)
    uniq = sorted(set(steps))
    # ★And the tolerance is COM's, not arithmetic's: it answers 3.1249 and
    # 3.1251 for the same eighth, so 1e-6 counts float noise as a violation.
    off = [s for s in uniq if abs(s * PER_PT - round(s * PER_PT)) > 0.02]
    print(f"  {face!r} {size}pt, {len(text)} characters"
          + (f", {wraps} line wrap(s) skipped" if wraps else ""))
    print(f"  distinct steps ({len(uniq)}): {uniq[:14]}")
    print(f"  NOT on the 1/{int(PER_PT)}pt grid: {len(off)} of {len(uniq)}  {off[:8]}")
    if args.chars:
        for k in range(min(args.chars, len(steps))):
            print(f"    {text[k]!r:4} -> {steps[k]:7.4f}")

    if args.pdf and pdf:
        got = pdf_line(pdf, args.slide, text)
        if not got:
            print("  the PDF does not carry this line as one run")
            return
        n = min(len(got), len(xs))
        a = [v - xs[0] for v in xs[:n]]
        b = [v - got[0] for v in got[:n]]
        w = max(range(n), key=lambda k: abs(a[k] - b[k]))
        print(f"  vs the truth PDF over {n} characters: worst {abs(a[w] - b[w]):.3f}pt "
              f"at {text[w]!r}, last {a[-1] - b[-1]:+.3f}pt")
        print("  ★that divergence is the PDF's, not the engine's -- it is the floor "
              "of any instrument that reads glyph positions out of the export")


if __name__ == "__main__":
    main()
