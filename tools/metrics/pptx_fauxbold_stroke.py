# -*- coding: utf-8 -*-
"""How much does PowerPoint thicken a bold it had to SYNTHESISE? Read it off the PDF.

S-FAUXBOLD (2026-08-26) measured the thickening off a raster -- ink run widths on
one scanline of one slide -- got `about 1.2% of em`, and the corpus gate washed
because a single constant does not generalise. The ledger's own next step was
"measure the stroke width per face and per size". This does that WITHOUT a
raster: when PowerPoint cannot find a bold face it emits the run in text
rendering mode 2 (fill THEN stroke) and states the pen width in the content
stream. That number is the thickening, exactly, with no threshold and no
antialiasing in the way ([[ink_is_not_a_weight_ruler]] -- black share measures
AA, not weight; this measures neither, it reads the instruction).

Reported per (deck, page, font, size): the stroke width in text space and its
ratio to the font size, so a constant-fraction law and a constant-absolute law
are told apart by looking.

The same scan answers the other half of "what did PowerPoint synthesise": a
faked ITALIC is a text matrix with a horizontal shear and no rotation, and its
factor is printed alongside.

Only the PAGE content stream is read. That is complete for these exports:
across all 1831 pages of both corpora, 28 pages carry a form XObject and NOT
ONE of them contains a `Tf` -- so no text, and no pen, hides in one. What CAN
hide is a page PowerPoint rasterised instead of setting (d32 p1's 223pt title
is an image, [[pptx_pdf_stencil_layer]]): there the pen is baked into pixels
and this reports nothing, which reads the same as "no synthesis here".

Usage:
    python tools/metrics/pptx_fauxbold_stroke.py [--corpus blind|dev|both]
                                                 [--deck 35] [--raw]
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
BLIND = REPO / "pipeline_data" / "pptx_benchmark" / "ssim_pptx" / "ppt_pdf"
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev" / "pdf"

# One PDF token: a name, a number, a string, an array/dict delimiter, or an
# operator. Strings are skipped wholesale so that a `(w)` inside text cannot be
# read as the line-width operator -- the mistake that makes naive scanners
# report line widths no page has.
TOKEN = re.compile(rb"""
      /(?P<name>[^\s/<>\[\]()]+)
    | (?P<num>[-+]?[0-9]*\.?[0-9]+)
    | (?P<str>\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]*>)
    | (?P<delim>[\[\]{}]|<<|>>)
    | (?P<op>[A-Za-z'"*]+)
""", re.X | re.S)


def mat_scale(m: tuple[float, ...]) -> float:
    """Uniform-ish scale of a 2x2: the geometric mean of the two axis lengths."""
    a, b, c, d = m[0], m[1], m[2], m[3]
    sx = (a * a + b * b) ** 0.5
    sy = (c * c + d * d) ** 0.5
    return (sx * sy) ** 0.5


def mul(m: tuple[float, ...], n: tuple[float, ...]) -> tuple[float, ...]:
    a, b, c, d, e, f = m
    A, B, C, D, E, F = n
    return (a * A + b * C, a * B + b * D,
            c * A + d * C, c * B + d * D,
            e * A + f * C + E, e * B + f * D + F)


def scan_page(content: bytes, out_shear: list | None = None) -> list[dict]:
    """Every text-showing op drawn in a stroking mode, with the pen that drew it.

    `out_shear` collects the sheared (faked-italic) ops, which the same pass
    sees for free."""
    if out_shear is None:
        out_shear = []
    ctm = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    stack: list[tuple] = []
    lw = 1.0
    tr = 0
    font = ""
    size = 0.0
    tm = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    ops: list[float] = []
    out: list[dict] = []
    shear: list[dict] = []
    for m in TOKEN.finditer(content):
        if m.group("num") is not None:
            ops.append(float(m.group("num")))
            continue
        if m.group("name") is not None:
            ops.append(m.group("name").decode("latin-1"))
            continue
        if m.group("str") is not None or m.group("delim") is not None:
            continue
        op = m.group("op").decode("latin-1")
        if op == "q":
            stack.append((ctm, lw, tr, font, size))
        elif op == "Q":
            if stack:
                # ★The render MODE is graphics state too. Restoring only the
                # matrix leaves `2 Tr` set after its q/Q block closes, and every
                # later fill-mode run is then reported as stroked -- with the
                # PDF's DEFAULT 1.0 pen, since no `w` was ever issued for it.
                # Those rows (ratio 1/size, never 1/35) were the tell.
                ctm, lw, tr, font, size = stack.pop()
        elif op == "cm" and len(ops) >= 6:
            ctm = mul(tuple(ops[-6:]), ctm)
        elif op == "w" and ops:
            lw = float(ops[-1])
        elif op == "Tf" and len(ops) >= 2:
            font, size = str(ops[-2]), float(ops[-1])
        elif op == "Tr" and ops:
            tr = int(ops[-1])
        elif op == "Tm" and len(ops) >= 6:
            tm = tuple(ops[-6:])
        elif op == "BT":
            tm = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        elif op in ("Tj", "TJ", "'", '"'):
            # Modes 2 and 6 fill AND stroke: that is the synthesised bold.
            if abs(tm[1]) < 1e-6 and abs(tm[2]) > 1e-6 and abs(tm[0] - tm[3]) < 1e-6:
                shear.append({"shear": tm[2] / tm[0] if tm[0] else 0.0,
                              "size": size * mat_scale(tm)})
            if tr in (1, 2, 5, 6):
                out.append({
                    "font": font,
                    "size": size * mat_scale(tm),
                    "raw_size": size,
                    "stroke": lw * mat_scale(ctm),
                    "raw_stroke": lw,
                    "ctm": mat_scale(ctm),
                    "tm": mat_scale(tm),
                    "mode": tr,
                })
        ops = []
    out_shear.extend(shear)
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", choices=("blind", "dev", "both"), default="both")
    ap.add_argument("--deck", default="")
    ap.add_argument("--raw", action="store_true", help="print every op, not the summary")
    args = ap.parse_args()

    roots = []
    if args.corpus in ("blind", "both"):
        roots.append(("blind", BLIND))
    if args.corpus in ("dev", "both"):
        roots.append(("dev", DEV))

    per_key: dict[tuple, list[dict]] = defaultdict(list)
    sheared: list[tuple[str, int, float, float]] = []
    for label, root in roots:
        for pdf_path in sorted(root.glob("*.pdf")):
            deck = pdf_path.stem.split("__")[0]
            if args.deck and args.deck not in deck:
                continue
            try:
                doc = pymupdf.open(pdf_path)
            except Exception as exc:
                print(f"  {deck}: {exc}")
                continue
            for pno in range(len(doc)):
                shear: list[dict] = []
                try:
                    hits = scan_page(doc[pno].read_contents(), shear)
                except Exception as exc:
                    print(f"  {label} {deck} p{pno + 1}: {exc}")
                    continue
                for h in shear:
                    sheared.append((f"{label}:{deck}", pno + 1, h["shear"], h["size"]))
                for h in hits:
                    h["deck"] = f"{label}:{deck}"
                    h["page"] = pno + 1
                    per_key[(h["deck"], h["font"], round(h["size"], 2))].append(h)
            doc.close()

    if sheared:
        print("faked ITALIC -- text matrices with a shear and no rotation")
        for deck, page, sh, size in sheared:
            print(f"  {deck:<10} p{page:<3} shear {sh:.4f}  size {size:.2f}"
                  f"  ({'1/3 exactly' if abs(sh - 1 / 3) < 5e-4 else 'other'})")
        print()
    if not per_key:
        print("no stroked text found -- no deck in this corpus synthesises bold")
        return

    if args.raw:
        for key in sorted(per_key):
            for h in per_key[key]:
                print(f"{h['deck']:>10} p{h['page']:<3} {h['font']:<10} "
                      f"size {h['size']:8.3f} stroke {h['stroke']:8.4f} "
                      f"ratio {h['stroke'] / h['size']:.5f} mode {h['mode']}")
        return

    print(f"{'deck':<10}{'font':<10}{'size':>9}{'stroke':>9}{'ratio':>9}"
          f"{'em%':>8}{'ops':>6}{'pages':>7}")
    for key in sorted(per_key, key=lambda k: (k[0], k[2])):
        rows = per_key[key]
        size = key[2]
        strokes = sorted({round(r["stroke"], 4) for r in rows})
        pages = len({r["page"] for r in rows})
        s = strokes[0]
        spread = "" if len(strokes) == 1 else f"  ({len(strokes)} distinct)"
        print(f"{key[0]:<10}{key[1]:<10}{size:>9.3f}{s:>9.4f}"
              f"{s / size if size else 0:>9.5f}{100 * s / size if size else 0:>7.2f}%"
              f"{len(rows):>6}{pages:>7}{spread}")


if __name__ == "__main__":
    main()
