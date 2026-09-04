# -*- coding: utf-8 -*-
"""How often does PowerPoint SHEAR an upright face instead of using an italic one?

d15 slide 5 asks for `b="1" i="1"` on Barlow, a family whose cloud cache holds a
real Bold Italic -- and PowerPoint drew the upright Bold with a text matrix of
`[1 0 0.3333 1 ...]`. Its advances are the Bold's (350.40pt for the trimmed
line) where the Bold Italic's are 341.97, so this is not a cosmetic difference:
it decides where the line breaks.

The truth PDFs state which it did, per span: a non-zero `c` in the text matrix
is a synthesised slant, and the font name beside it says which face was used.
Beside each one, WHERE that family lives -- the deck's own embedded parts, the
Office cloud cache, or the machine -- because that is the candidate
discriminator and the census has to be able to falsify it.

    python tools/metrics/pptx_shear_census.py            # the dev corpus
    python tools/metrics/pptx_shear_census.py --blind
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

import pymupdf
from fontTools.ttLib import TTFont

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2] / "pipeline_data" / "pptx_benchmark"

# `/F1 30 Tf` selects a font; `a b c d e f Tm` sets the text matrix. Reading
# them in stream order pairs each matrix with the face in force.
TOKEN = re.compile(
    rb"/(\w+)\s+[\d.]+\s+Tf"
    rb"|(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+Tm"
)
ITALIC_NAME = re.compile(r"italic|oblique", re.I)
CLOUD_ROOT = Path(os.path.expandvars(
    r"%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts"))


def page_rows(page) -> list[tuple[str, float]]:
    """(font name, shear) for every text matrix on the page."""
    fonts = {f[4]: f[3] for f in page.get_fonts()}
    out = []
    face = None
    for m in TOKEN.finditer(page.read_contents()):
        if m.group(1):
            face = fonts.get(m.group(1).decode("latin-1"), "?")
        elif face is not None:
            b, c = float(m.group(3)), float(m.group(4))
            # `b` is rotation, `c` is the slant PowerPoint adds itself.
            out.append((face, 0.0 if abs(b) > 1e-6 else c))
    return out


def cloud_families() -> set[str]:
    """Families the Office cloud cache holds, by their own sfnt name."""
    out = set()
    for path in CLOUD_ROOT.rglob("*"):
        if path.suffix.lower() not in (".ttf", ".otf", ".ttc"):
            continue
        try:
            font = TTFont(str(path), lazy=True, fontNumber=0)
            out.add(font["name"].getDebugName(1))
        except Exception:
            continue
    return {f for f in out if f}


def norm(name: str) -> str:
    return "".join(c for c in name.lower() if c.isalnum())


def where(face: str, cloud: set[str], embedded: set[str]) -> str:
    """cloud / embedded / installed, for the family a PDF font name carries."""
    # `BCDEEE+Barlow,Bold` -> `Barlow`; `Arial-BoldItalicMT` -> `Arial`.
    fam = face.split("+")[-1].split(",")[0].split("-")[0]
    # The cloud copy is asked FIRST: a family the cache holds is served from
    # there even when the deck also ships parts for it, so testing `embedded`
    # first would file every cloud face under the wrong home.
    if norm(fam) in {norm(c) for c in cloud}:
        return "cloud"
    # A deck that renames its parts numerically subsets them under the number.
    if fam.isdigit() or norm(fam) in {norm(e) for e in embedded}:
        return "embedded"
    return "installed"


def embedded_typefaces(pptx: Path | None) -> set[str]:
    """The families a deck ships parts for, from `p:embeddedFont`."""
    if not pptx or not pptx.exists():
        return set()
    try:
        with zipfile.ZipFile(pptx) as z:
            xml = z.read("ppt/presentation.xml").decode("utf-8", "replace")
    except Exception:
        return set()
    return set(re.findall(r'<p:font typeface="([^"]+)"', xml))


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--blind", action="store_true")
    args = ap.parse_args()
    sub = "ssim_pptx/ppt_pdf" if args.blind else "dev/pdf"
    src = ROOT / ("ssim_pptx/pptx" if args.blind else "dev/pptx")
    pdfs = sorted((ROOT / sub).glob("*.pdf"))

    cloud = cloud_families()
    tally: Counter = Counter()
    for path in pdfs:
        hits = sorted(src.glob(path.stem + "*.pptx"))
        embedded = embedded_typefaces(hits[0] if hits else None)
        faces: Counter = Counter()
        with pymupdf.open(path) as doc:
            for page in doc:
                for face, shear in page_rows(page):
                    real = bool(ITALIC_NAME.search(face))
                    if abs(shear) < 1e-6 and not real:
                        continue
                    kind = "SYNTH" if abs(shear) > 1e-6 else "REAL"
                    faces[(kind, where(face, cloud, embedded),
                           face.split("+")[-1])] += 1
        if faces:
            print("%s" % path.stem[:44])
            for (kind, home, face), n in sorted(faces.items()):
                print("      %-5s %-9s %-30s x%d" % (kind, home, face, n))
        tally.update({(k, h): n for (k, h, _f), n in faces.items()})
    print("\n%d decks -- italic, by where the face lives:" % len(pdfs))
    for (kind, home), n in sorted(tally.items()):
        print("   %-5s %-9s %d spans" % (kind, home, n))
    print("SYNTH = PowerPoint sheared an upright face rather than use an italic one.")


if __name__ == "__main__":
    main()
