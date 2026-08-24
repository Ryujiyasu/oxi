"""Where does Excel start a shape's text, for a stated left inset?

`b6a3a84180c9_002` — the corpus's lowest-scoring workbook — carries a text box
whose `<a:bodyPr>` states `lIns="288000"`, and every glyph of its two visible
lines sits exactly 3px right of where Oxi puts it. The whole line is a
translation, not a stretch: the first blob and the last are both off by the
same 3. The box itself is in the right place — the border's right edge lands
on the same pixel in both renders — so the miss is between the box's edge and
the first letter.

That book cannot answer the question, because its box starts off the left of
the captured picture and its own left edge cannot be read. So the question is
put to Excel on a box that sits entirely inside the sheet: sweep the stated
inset, and read how far the ink starts from the fill's own left edge.

`lIns` is not reachable through COM's shape API in the form needed, so the
values are written into the drawing XML by hand.

Run: python tools/metrics/_xlsx_shape_inset.py
"""

from __future__ import annotations

import os
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
SCRATCH = Path(r"C:\tmp\xlsx_shape_inset")
# EMU per pixel at 96 dpi.
EMU = 9525
INSETS = [0, 45720, 91440, 182880, 288000, 457200]
# Which edge each block of shapes varies: the same conversion serves all
# four, so fixing one without asking about the others would be a guess.
EDGES = ["l", "t", "r"]
WORDS = "あいうえお"
FILL = "DEEBF7"


def seed() -> Path:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "seed.xlsx"
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:P90").Interior.Color = 0xFFFFFF
        for at in range(len(INSETS) * len(EDGES)):
            shape = sheet.Shapes.AddShape(1, 40.0, 20.0 + at * 60.0, 420.0, 44.0)
            shape.TextFrame2.TextRange.Text = WORDS
            shape.TextFrame2.TextRange.Font.Size = 14
            shape.TextFrame2.TextRange.Font.Name = "ＭＳ ゴシック"
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = (
                3 if EDGES[at // len(INSETS)] == 'r' else 1)
            # The default shape style writes its text in white, which is
            # invisible on the light fill this uses.
            shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x000000
            shape.Fill.ForeColor.RGB = 0xF7EBDE     # BGR of DEEBF7
            shape.Line.Visible = False
        book.SaveAs(str(made), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return made


def written(source: Path) -> Path:
    out = SCRATCH / "insets.xlsx"
    if out.exists():
        out.unlink()
    held = zipfile.ZipFile(source)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as writing:
        for item in held.infolist():
            raw = held.read(item.filename)
            if item.filename.startswith("xl/drawings/") and item.filename.endswith(".xml"):
                text = raw.decode("utf-8")
                bodies = list(re.finditer(r"<a:bodyPr[^>]*/>|<a:bodyPr[^>]*>", text))
                arms = [(edge, inset) for edge in EDGES for inset in INSETS]
                assert len(bodies) >= len(arms), f"{len(bodies)} bodyPr for {len(arms)} arms"
                for spot, (edge, inset) in zip(reversed(bodies[: len(arms)]), reversed(arms)):
                    sides = {"l": 91440, "t": 45720, "r": 91440}
                    sides[edge] = inset
                    fresh = (f'<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" '
                             f'lIns="{sides["l"]}" tIns="{sides["t"]}" rIns="{sides["r"]}" '
                             f'bIns="45720" anchor="t"/>')
                    text = text[: spot.start()] + fresh + text[spot.end():]
                raw = text.encode("utf-8")
            writing.writestr(item, raw)
    return out


def read(grey, rgb_mask, y0, y1, edge):
    """How far the ink sits from the box's own edge, on the edge being swept."""
    fill_x = np.where(rgb_mask[y0:y1].any(axis=0))[0]
    fill_y = np.where(rgb_mask[y0:y1].any(axis=1))[0]
    lit = (grey[y0:y1] < 140)
    ink_x = np.where(lit.any(axis=0))[0]
    ink_y = np.where(lit.any(axis=1))[0]
    if not len(fill_x) or not len(ink_x):
        return None
    if edge == "l":
        return int(ink_x.min()) - int(fill_x.min())
    if edge == "r":
        return int(fill_x.max()) - int(ink_x.max())
    return int(ink_y.min()) - int(fill_y.min())


def main() -> int:
    book_path = written(seed())
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(book_path))
    tops = []
    try:
        sheet = book.Worksheets(1)
        for at in range(len(INSETS) * len(EDGES)):
            shape = sheet.Shapes(at + 1)
            tops.append((round(shape.Top * 96 / 72), round(shape.Height * 96 / 72)))
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:P90").CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(0.7)
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

    out = {}
    for name, path in (("excel", SCRATCH / "excel.png"), ("ours", SCRATCH / "oxi.png")):
        rgb = np.asarray(Image.open(path).convert("RGB")).astype(int)
        grey = np.asarray(Image.open(path).convert("L")).astype(int)
        mask = ((rgb[:, :, 0] == 0xDE) & (rgb[:, :, 1] == 0xEB) & (rgb[:, :, 2] == 0xF7))
        out[name] = [read(grey, mask, top, top + high, EDGES[at // len(INSETS)])
                     for at, (top, high) in enumerate(tops)]

    import math
    print(f"  a {FILL} box, {WORDS} — how far the ink sits from the box's own edge")
    print("  edge  inset EMU      px   Excel  Oxi  delta   round  ceil")
    for at, (edge, inset) in enumerate([(e, i) for e in EDGES for i in INSETS]):
        egap, ogap = out["excel"][at], out["ours"][at]
        delta = None if (egap is None or ogap is None) else ogap - egap
        flag = "" if delta == 0 else "   <--"
        px = inset / EMU
        print(f"  {edge:<5} {inset:>9}  {px:>7.2f}   {str(egap):>5} {str(ogap):>4}"
              f"   {'' if delta is None else f'{delta:+d}':>5}"
              f"   {round(px):>5} {math.ceil(px):>5}{flag}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
