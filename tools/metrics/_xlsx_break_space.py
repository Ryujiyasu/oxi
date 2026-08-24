"""Which spaces does Excel leave behind at a wrap?

`broken_at` already knows that a space a line breaks on stays with the line it
ends, so the next line starts at the first character past it. It only knows
that about the ASCII space. `13ea087fa546_data_A22` wraps a 791-character
merged cell whose separators are all U+3000, and Excel starts the new line at
the flush margin while Oxi indents it by one full-width space — 12px, the
book's worst tile.

One document is not a rule. This asks Excel directly: a wrapped cell of a
fixed width, and a text whose break falls exactly on a space, for each kind of
space worth asking about. What is read is where the ink of the SECOND line
starts, against where the first line's starts. Equal means the space was left
behind; indented means it was carried down.

Run: python tools/metrics/_xlsx_break_space.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_break_space")
FACE = "ＭＳ 明朝"
SIZE = 11.0
# Each arm: what the space is, and a name for it.
SPACES = [
    (" ", "ASCII space U+0020"),
    ("\u3000", "ideographic space U+3000"),
    ("\u00a0", "no-break space U+00A0"),
    ("  ", "two ASCII spaces"),
    ("\u3000\u3000", "two ideographic spaces"),
]
HEAD = "あいうえおかきくけこさしすせそ"
TAIL = "たちつてとなにぬねの"


def picture(sheet):
    for _ in range(8):
        try:
            sheet.Activate()
            sheet.Range("A1:H10").CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.4)
        held = ImageGrab.grabclipboard()
        if held is not None:
            return held
    return None


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:H10").Interior.Color = 0xFFFFFF
        cell = sheet.Range("B2")
        sheet.Columns("B").ColumnWidth = 16.0
        sheet.Rows("2").RowHeight = 60.0
        cell.WrapText = True
        cell.HorizontalAlignment = -4131      # xlLeft
        cell.VerticalAlignment = -4160        # xlTop
        cell.Font.Name = FACE
        cell.Font.Size = SIZE
        print(f"  {FACE} {SIZE:.0f}pt in a 16-wide wrapped cell")
        print("  what the break lands on            line 1 x   line 2 x   verdict")
        for space, name in SPACES:
            cell.Value = HEAD + space + TAIL
            held = picture(sheet)
            if held is None:
                print(f"   {name}: no picture")
                continue
            held.save(SCRATCH / f"{name.replace(' ', '_')}.png")
            top = round(cell.Top * 96 / 72)
            left = round(cell.Left * 96 / 72)
            grey = np.asarray(held.convert("L"))[
                top:top + round(cell.Height * 96 / 72),
                left:left + round(cell.Width * 96 / 72)]
            ink = grey < 140
            # Split into rows of ink, then take the first two.
            lit = ink.sum(axis=1) > 0
            runs, start = [], None
            for y, on in enumerate(lit):
                if on and start is None:
                    start = y
                elif not on and start is not None:
                    runs.append((start, y))
                    start = None
            if start is not None:
                runs.append((start, len(lit)))
            if len(runs) < 2:
                print(f"   {name:<32} only {len(runs)} line(s) of ink")
                continue
            firsts = []
            for y0, y1 in runs[:2]:
                x = np.where(ink[y0:y1].any(axis=0))[0]
                firsts.append(int(x.min()) if len(x) else -1)
            gap = firsts[1] - firsts[0]
            verdict = "left behind" if abs(gap) <= 1 else f"carried down ({gap:+d}px)"
            print(f"   {name:<32} {firsts[0]:>6}     {firsts[1]:>6}     {verdict}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
