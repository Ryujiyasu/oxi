"""Does a cell's NOTE clip its text to its box?

`b6a3a84180c9_002` is the corpus's lowest-scoring workbook, and its worst
tiles are all in one column band: the yellow note boxes down the right-hand
side. Excel draws three lines in a box and stops; Oxi draws a fourth, running
past where the box ends.

A shape drops the lines that do not fit when its body says
`vertOverflow="clip"` (SX61). A note carries no body properties at all — its
box comes from the VML — so nothing in the file says whether it clips. This
asks Excel.

The arm: one note, its box fixed, and more and more text put in it. What is
counted is how many lines of ink come out, against how many the text needs.

Run: python tools/metrics/_xlsx_note_clip.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_note_clip")
FACE = "ＭＳ ゴシック"
SIZE = 12.0
LINE = "いろはにほへとちりぬるを"
HIGHS = [40.0, 60.0, 80.0]
COUNTS = [2, 4, 6, 8]


def picture(sheet):
    for _ in range(8):
        try:
            sheet.Activate()
            sheet.Range("A1:Z40").CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.5)
        held = ImageGrab.grabclipboard()
        if held is not None:
            return held
    return None


def lines_of_ink(image, top: int, left: int, high: int, wide: int) -> int:
    grey = np.asarray(image.convert("L"))[top - 4:top + high + 30, left + 4:left + wide - 4]
    ink = grey < 120
    lit = ink.sum(axis=1) > 1
    count, was = 0, False
    for on in lit:
        if on and not was:
            count += 1
        was = on
    return count


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z40").Interior.Color = 0xFFFFFF
        target = sheet.Range("C5")
        print(f"  {FACE} {SIZE:.0f}pt; the box is held while the text grows")
        for high_pt in HIGHS:
            for count in COUNTS:
                if target.Comment is not None:
                    target.Comment.Delete()
                words = "\n".join([LINE] * count)
                note = target.AddComment(words)
                try:
                    note.Visible = True
                    frame = note.Shape.TextFrame
                    frame.Characters().Text = words
                    frame.Characters().Font.Name = FACE
                    frame.Characters().Font.Size = SIZE
                    frame.AutoSize = False
                    note.Shape.Width = 260.0
                    note.Shape.Height = high_pt
                    note.Shape.Fill.ForeColor.RGB = 0xFFFFFF
                    note.Shape.Line.Visible = False
                    held = picture(sheet)
                    if held is None:
                        print(f"    {high_pt:.0f}pt box, {count} lines: no picture")
                        continue
                    box = note.Shape
                    drawn = lines_of_ink(
                        held,
                        round(box.Top * 96 / 72),
                        round(box.Left * 96 / 72),
                        round(box.Height * 96 / 72),
                        round(box.Width * 96 / 72),
                    )
                    fits = int(box.Height * 96 / 72) // 16
                    print(f"    box {high_pt:>5.0f}pt ({round(high_pt*96/72):>3}px)"
                          f"  text {count} lines  ->  {drawn} drawn"
                          f"   (about {fits} would fit)")
                finally:
                    if target.Comment is not None:
                        target.Comment.Delete()
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
