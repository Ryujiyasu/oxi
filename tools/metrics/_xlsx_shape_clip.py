"""Does a shape drop the lines that do not fit whatever its `vertOverflow` says?

`_xlsx_shape_overflow.py` found that Excel draws only the lines that fit a
shape's text rectangle and anchors those — but every shape it measured was one
Excel had added itself, and Excel writes `vertOverflow="clip"` on those. The
corpus is not all clip: 256 shapes say clip, 322 say nothing at all (where
DrawingML's default is `overflow`), and 7 say overflow outright.

So the same box is drawn three times over, side by side, with the three
settings written into the drawing by hand. If all three drop the same lines
the rule is about the box; if only the clipped one drops them the rule has to
ask the body what it says first.

Run: python tools/metrics/_xlsx_shape_clip.py
"""

from __future__ import annotations

import re
import shutil
import sys
import time
import zipfile
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_shape_clip")
TEXT = ["いろはにほへと", "ちりぬるを", "わかよたれそ", "つねならむ", "うゐのおくやま"]
FACE = "ＭＳ ゴシック"
SIZE = 12.0
TOP_PT = 100.0
WIDE_PT = 150.0
HIGH_PT = 90.0  # holds three of the five lines
LEFTS = (60.0, 230.0, 400.0)
SAYS = ('vertOverflow="clip" ', "", 'vertOverflow="overflow" ')
NAMES = ("clip", "absent", "overflow")


def author() -> Path:
    """One workbook, three shapes, written by Excel then patched by hand."""
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "made.xlsx"
    if made.exists():
        made.unlink()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        for left in LEFTS:
            shape = sheet.Shapes.AddShape(1, left, TOP_PT, WIDE_PT, HIGH_PT)  # 1 = rect
            frame = shape.TextFrame2
            frame.WordWrap = True
            frame.AutoSize = 0
            frame.VerticalAnchor = 1  # top, so a dropped line shows as a missing foot
            frame.TextRange.Text = "\n".join(TEXT)
            frame.TextRange.Font.Size = SIZE
            frame.TextRange.Font.Name = FACE
            frame.TextRange.Font.Fill.ForeColor.RGB = 0
            shape.Fill.Visible = False
            shape.Line.Visible = False
        book.SaveAs(str(made), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()

    patched = SCRATCH / "patched.xlsx"
    if patched.exists():
        patched.unlink()
    held = zipfile.ZipFile(made)
    with zipfile.ZipFile(patched, "w", zipfile.ZIP_DEFLATED) as writing:
        for item in held.infolist():
            raw = held.read(item.filename)
            if item.filename.startswith("xl/drawings/") and item.filename.endswith(".xml"):
                text = raw.decode("utf-8")
                bodies = list(re.finditer(r"<a:bodyPr[^>]*/?>", text))
                assert len(bodies) == len(SAYS), f"{len(bodies)} bodies, wanted {len(SAYS)}"
                for body, said in zip(reversed(bodies), reversed(SAYS)):
                    fresh = re.sub(r'vertOverflow="[^"]*"\s*', "", body.group(0))
                    fresh = fresh.replace("<a:bodyPr ", f"<a:bodyPr {said}", 1)
                    text = text[: body.start()] + fresh + text[body.end() :]
                raw = text.encode("utf-8")
            writing.writestr(item, raw)
    return patched


def bands(image, left: int, right: int) -> list[tuple[int, int]]:
    grey = np.asarray(image.convert("L"))[:, left:right]
    lit = (grey < 120).sum(axis=1)
    out: list[tuple[int, int]] = []
    start = None
    for row, count in enumerate(lit):
        if count > 1 and start is None:
            start = row
        if count <= 1 and start is not None:
            out.append((start, row - 1))
            start = None
    if start is not None:
        out.append((start, len(lit) - 1))
    return out


def main() -> int:
    book_path = author()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(book_path))
    try:
        sheet = book.Worksheets(1)
        held = None
        for _ in range(6):
            try:
                sheet.Activate()
                sheet.Range("A1:Z60").CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.5)
            held = ImageGrab.grabclipboard()
            if held is not None:
                break
        if held is None:
            print("Excel would not hand over a picture")
            return 1
        held.save(SCRATCH / "shot.png")
        top = round(TOP_PT * 96 / 72)
        print(f"box top {top}, {round(HIGH_PT * 96 / 72)} tall, five lines of {SIZE:.0f}pt")
        for name, left in zip(NAMES, LEFTS):
            window = (round(left * 96 / 72) + 6, round((left + WIDE_PT) * 96 / 72) - 6)
            found = bands(held, *window)
            print(f"  {name:<9} n {len(found)} tops {[band[0] for band in found]}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
