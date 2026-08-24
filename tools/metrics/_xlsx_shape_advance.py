"""How wide does Excel make a shape's line, against how wide Oxi makes it?

Right-aligned text in a shape is a ruler for the line's total advance. The box
edge and the inset are known, both renderers agree on where the area's right
edge is when `rIns` is zero, and the glyphs are the same — so whatever gap is
left between the box and the last letter's ink is the width each engine
believes the line has, read off directly.

`_xlsx_shape_inset.py` found that gap to be a fixed 2 pixels in Excel and 0 in
Oxi, which says Oxi's line is two pixels short. Two pixels is either a
per-character rounding that grows with the line, or something fixed at its
end. This sweeps the length and the size to tell those apart: a per-character
error rises with the count, a trailing one does not.

The same sweep is the thing that has to land beside the `floor` on a shape's
far edge — correcting that edge alone walked 16 workbooks backwards, because
rounding the inset was cancelling this.

Run: python tools/metrics/_xlsx_shape_advance.py
"""

from __future__ import annotations

import re
import subprocess
import sys
import time
import zipfile
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_shape_advance")
# The law was first read off ONE glyph — a full-width あ, whose advance is the
# em exactly — and rounding each of those to a whole pixel matched Excel's ink
# width but cost the corpus 12 workbooks against 6. A glyph whose advance is
# NOT the em has to be in the sweep before any rule is believed.
LETTERS = ["あ", "W", "i", "1", "ｱ"]
COUNTS = [1, 8]
SIZES = [11.0, 20.0]
FACES = ["ＭＳ ゴシック", "ＭＳ Ｐゴシック"]


def seed(arms) -> Path:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "seed.xlsx"
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:R120").Interior.Color = 0xFFFFFF
        for at, (face, size, letter, count) in enumerate(arms):
            shape = sheet.Shapes.AddShape(1, 40.0, 14.0 + at * 44.0, 460.0, 34.0)
            shape.TextFrame2.TextRange.Text = letter * count
            shape.TextFrame2.TextRange.Font.Size = size
            shape.TextFrame2.TextRange.Font.Name = face
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 3     # right
            shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x000000
            shape.Fill.ForeColor.RGB = 0xF7EBDE
            shape.Line.Visible = False
        book.SaveAs(str(made), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return made


def written(source: Path, arms) -> Path:
    """Every box gets a zero right inset, so the area's edge is the box's."""
    out = SCRATCH / "advance.xlsx"
    if out.exists():
        out.unlink()
    held = zipfile.ZipFile(source)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as writing:
        for item in held.infolist():
            raw = held.read(item.filename)
            if item.filename.startswith("xl/drawings/") and item.filename.endswith(".xml"):
                text = raw.decode("utf-8")
                bodies = list(re.finditer(r"<a:bodyPr[^>]*/>|<a:bodyPr[^>]*>", text))
                assert len(bodies) >= len(arms), f"{len(bodies)} bodyPr for {len(arms)} arms"
                for spot in reversed(bodies[: len(arms)]):
                    fresh = ('<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" '
                             'lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/>')
                    text = text[: spot.start()] + fresh + text[spot.end():]
                raw = text.encode("utf-8")
            writing.writestr(item, raw)
    return out


def main() -> int:
    arms = [(face, size, letter, count)
            for face in FACES for size in SIZES
            for letter in LETTERS for count in COUNTS]
    book_path = written(seed(arms), arms)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(book_path))
    boxes = []
    try:
        sheet = book.Worksheets(1)
        for at in range(len(arms)):
            shape = sheet.Shapes(at + 1)
            boxes.append((round(shape.Top * 96 / 72), round(shape.Height * 96 / 72)))
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:R120").CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(0.8)
            shot = ImageGrab.grabclipboard()
            if shot is not None:
                break
        else:
            print("Excel would not hand over a picture")
            return 1
        shot.save(SCRATCH / "excel.png")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()

    subprocess.run([str(RENDERER), str(book_path), str(SCRATCH / "oxi.png")],
                   capture_output=True, check=False)

    def gaps(path):
        rgb = np.asarray(Image.open(path).convert("RGB")).astype(int)
        grey = np.asarray(Image.open(path).convert("L")).astype(int)
        fill = ((rgb[:, :, 0] == 0xDE) & (rgb[:, :, 1] == 0xEB) & (rgb[:, :, 2] == 0xF7))
        out = []
        for top, high in boxes:
            band_fill = np.where(fill[top:top + high].any(axis=0))[0]
            band_ink = np.where((grey[top:top + high] < 140).any(axis=0))[0]
            out.append((int(band_fill.max()) - int(band_ink.max()),
                        int(band_ink.max()) - int(band_ink.min()) + 1)
                       if len(band_fill) and len(band_ink) else (None, None))
        return out

    theirs, ours = gaps(SCRATCH / "excel.png"), gaps(SCRATCH / "oxi.png")
    # The advance each engine believes, solved from the ink's width:
    # span(8) - span(1) = 7 * advance.
    print("  right-aligned, box with no insets; advance solved from the ink width")
    print("  face          size letter   Excel adv   Oxi adv   exact em*share")
    spans = {}
    for at, (face, size, letter, count) in enumerate(arms):
        spans[(face, size, letter, count)] = (theirs[at][1], ours[at][1])
    for face in FACES:
        for size in SIZES:
            for letter in LETTERS:
                one, many = spans[(face, size, letter, 1)], spans[(face, size, letter, 8)]
                if None in one or None in many:
                    print(f"  {face:<13}{size:>4.0f} {letter:<6}   no ink")
                    continue
                e_adv = (many[0] - one[0]) / 7
                o_adv = (many[1] - one[1]) / 7
                flag = "" if abs(e_adv - o_adv) < 0.08 else "   <--"
                print(f"  {face:<13}{size:>4.0f} {letter:<6}   {e_adv:>9.3f}   {o_adv:>7.3f}"
                      f"{flag}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
