"""Which face did PowerPoint actually set this line in?

The truth PDF places every glyph, so the advances it uses ARE the face's -- and
a face can be identified by them. Given a line, this walks every font file on
the machine and reports which one advances the way the PDF does.

★It exists because the obvious check is not enough. d15 slide 5's quotation was
compared against the family the layout believed it was in (Barlow Bold), matched
to 0.000%, and the difference was therefore blamed on something else -- eight
hypotheses, all dead. The face was Barlow Bold ITALIC, asked for by the layout's
level and never by a run. Comparing against ONE face answers "is it this one",
which is not the question; the question is "which one is it".

    python tools/metrics/pptx_face_identify.py --deck d29 --slide 6 --text Mission
"""
from __future__ import annotations

import argparse
import glob
import io
import os
import sys
from pathlib import Path

import pymupdf
from fontTools.ttLib import TTFont

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOTS = [
    r"C:\Windows\Fonts",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "FontCache",
                 "4", "CloudFonts"),
]


def placed_advances(pdf: Path, page_no: int, prefix: str):
    """The advances the PDF actually stepped by, in EM, plus what it says it used."""
    doc = pymupdf.open(pdf)
    page = doc[page_no - 1]
    for span in page.get_texttrace():
        text = "".join(chr(c[0]) for c in span["chars"] if c[0] < 0x110000)
        if not text.startswith(prefix):
            continue
        size = span["size"]
        origins = [c[2][0] for c in span["chars"]]
        adv = {}
        for i in range(len(origins) - 1):
            ch = chr(span["chars"][i][0])
            adv.setdefault(ch, round((origins[i + 1] - origins[i]) / size, 4))
        doc.close()
        return span["font"], size, text, adv
    doc.close()
    return None, None, None, None


def faces():
    """Every font file this machine can offer, with its identity."""
    seen = set()
    for root in ROOTS:
        for path in glob.glob(os.path.join(root, "**", "*.tt[fc]"), recursive=True) + \
                    glob.glob(os.path.join(root, "**", "*.otf"), recursive=True):
            if path.lower() in seen:
                continue
            seen.add(path.lower())
            try:
                font = TTFont(path, lazy=True, fontNumber=0)
                name = font["name"]
                yield (name.getDebugName(6) or "",
                       name.getDebugName(1) or "",
                       name.getDebugName(2) or "",
                       os.path.basename(path), font)
            except Exception:
                continue


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", required=True)
    ap.add_argument("--slide", type=int, required=True)
    ap.add_argument("--text", required=True, help="the start of the line")
    ap.add_argument("--family", default="", help="only faces whose family contains this")
    args = ap.parse_args()

    pdf = next(iter(sorted(
        (REPO / "pipeline_data/pptx_benchmark/dev/pdf").glob(args.deck + "*.pdf"))), None)
    if not pdf:
        pdf = next(iter(sorted(
            (REPO / "pipeline_data/pptx_benchmark/ssim_pptx/ppt_pdf").glob(
                args.deck + "*.pdf"))), None)
    if not pdf:
        sys.exit(f"no truth PDF for {args.deck}")

    declared, size, text, adv = placed_advances(pdf, args.slide, args.text)
    if not adv:
        sys.exit(f"no line starting {args.text!r} on page {args.slide}")
    print(f"{args.deck} s{args.slide}: {text[:46]!r}")
    print(f"   the PDF says it used {declared!r} at {size:.3f}pt")
    print(f"   and stepped by: " +
          " ".join(f"{c!r}:{v}" for c, v in list(adv.items())[:8]))

    rows = []
    for ps, family, sub, filename, font in faces():
        if args.family and args.family.lower() not in family.lower():
            continue
        try:
            upm = font["head"].unitsPerEm
            cmap = font.getBestCmap()
            hmtx = font["hmtx"]
            if not cmap or not upm:
                continue
        except Exception:
            continue
        hits, total, worst = 0, 0, 0.0
        for ch, want in adv.items():
            g = cmap.get(ord(ch))
            if not g or g not in hmtx.metrics:
                continue
            total += 1
            got = hmtx[g][0] / upm
            worst = max(worst, abs(got - want))
            if abs(got - want) < 0.003:
                hits += 1
        if total >= max(3, len(adv) // 2):
            rows.append((hits / total, hits, total, worst, ps or f"{family} {sub}",
                         filename))
    rows.sort(key=lambda r: (-r[0], r[3]))
    print(f"\n{'match':>7} {'worst':>8}  {'postscript name':30} file")
    for frac, hits, total, worst, ps, filename in rows[:8]:
        print(f"{hits:3}/{total:<3} {worst:8.4f}  {ps[:29]:30} {filename}")
    if rows and rows[0][0] < 1.0:
        print("\nNo face advances exactly the way the PDF stepped -- the line may "
              "carry tracking, or be set in a face this machine does not have.")


if __name__ == "__main__":
    main()
