# -*- coding: utf-8 -*-
"""Probe: what does `a:srcRect` do to a shape's picture fill in PowerPoint?

d30 slide 16 (SSIM 0.5236, one of the corpus's worst) has a full-width photo
whose fill declares `<a:srcRect b="14368"/>`. Correlating the two renders band
by band shows the horizontal scale matches to 1.000 while the VERTICAL content
of Oxi is ~1.15x larger -- i.e. exactly the sort of difference a mis-modelled
source crop makes. The crop rule Oxi implements was derived on the DOCX side
(a Word `p:pic`), never measured for a PPTX shape blipFill.

Source is the 2x2 colour grid (TL red, TR green, BL blue, BR yellow) with a
black frame, so each sample point names the source quadrant that landed there.

Arms (4in square shape, centred):
  F1 no srcRect                   control
  F2 srcRect b="50000"            keep the TOP half of the source
  F3 srcRect t="50000"            keep the BOTTOM half
  F4 srcRect l="50000"            keep the RIGHT half
  F5 srcRect b="14368"            the literal d30 value
  F6 srcRect t="25000" b="25000"  keep the middle band
"""
from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Emu

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(r"pipeline_data\pptx_probes\srcrect").resolve()
SHAPE_EMU = 3657600
SHAPE_XY = (2743200, 1600200)

ARMS = [
    ("F1 no srcRect", ""),
    ("F2 srcRect b=50%", 'b="50000"'),
    ("F3 srcRect t=50%", 't="50000"'),
    ("F4 srcRect l=50%", 'l="50000"'),
    ("F5 srcRect b=14.368% (d30)", 'b="14368"'),
    ("F6 srcRect t=25% b=25%", 't="25000" b="25000"'),
]


def grid_png() -> bytes:
    img = Image.new("RGB", (400, 400), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 199, 199], fill=(255, 0, 0))
    d.rectangle([200, 0, 399, 199], fill=(0, 200, 0))
    d.rectangle([0, 200, 199, 399], fill=(0, 0, 255))
    d.rectangle([200, 200, 399, 399], fill=(255, 220, 0))
    # A thin black rule every 10% of the height, so a vertical crop is readable
    # to better than a quadrant.
    for i in range(1, 10):
        y = int(400 * i / 10)
        d.line([(0, y), (399, y)], fill=(0, 0, 0), width=3)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    tmp, dst = OUT / "_stage.pptx", OUT / "srcrect.pptx"
    png = grid_png()
    (OUT / "source.png").write_bytes(png)

    prs = Presentation()
    blank = prs.slide_layouts[6]
    for label, _ in ARMS:
        s = prs.slides.add_slide(blank)
        box = s.shapes.add_textbox(Emu(228600), Emu(228600), Emu(6400800), Emu(400050))
        box.text_frame.text = label
        s.shapes.add_picture(io.BytesIO(png), Emu(0), Emu(0), Emu(91440), Emu(91440))
        sp = s.shapes.add_shape(1, Emu(SHAPE_XY[0]), Emu(SHAPE_XY[1]),
                                Emu(SHAPE_EMU), Emu(SHAPE_EMU))
        sp.line.fill.background()
        sp.shadow.inherit = False
    prs.save(tmp)

    with zipfile.ZipFile(tmp) as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}

    for i, (label, src) in enumerate(ARMS, start=1):
        part = f"ppt/slides/slide{i}.xml"
        xml = data[part].decode("utf-8")
        rels = data[f"ppt/slides/_rels/slide{i}.xml.rels"].decode("utf-8")
        rid = next(t.split('"')[0] for t in rels.split('Id="')[1:] if "/image" in t)
        ps = xml.index("<p:pic>")
        pe = xml.index("</p:pic>") + len("</p:pic>")
        xml = xml[:ps] + xml[pe:]
        src_el = f"<a:srcRect {src}/>" if src else ""
        fill = (
            f'<a:blipFill rotWithShape="1"><a:blip r:embed="{rid}"><a:alphaModFix/></a:blip>'
            f"{src_el}<a:stretch><a:fillRect/></a:stretch></a:blipFill>"
        )
        # `ge` is already the index PAST the closing tag; adding the length a
        # second time spliced the fill into the middle of <a:ln><a:noFill/>.
        ge = xml.rindex("</a:prstGeom>") + len("</a:prstGeom>")
        xml = xml[:ge] + fill + xml[ge:]
        data[part] = xml.encode("utf-8")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])
    tmp.unlink()
    print(f"wrote {dst}  ({len(ARMS)} arms)")


if __name__ == "__main__":
    main()
