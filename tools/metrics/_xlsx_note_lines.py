# -*- coding: utf-8 -*-
r"""Where does the first line of a note's text sit, and how far apart are the rest?

With `002`'s note box now ending where Excel ends it, what is left inside it is
the text: our first line stands two pixels above Excel's and the gap to the
second is one pixel wider. Everything else in that corner of the sheet agrees,
so those three pixels are most of the workbook's remaining difference.

The box is a known quantity — `_xlsx_note_box.py` settled its height — so this
asks only about what is inside it: how far the first line's ink starts below
the box's own top edge, and how far apart the lines run, over several sizes and
faces and several counts of lines.

    python tools\metrics\_xlsx_note_lines.py
    python tools\metrics\_xlsx_note_lines.py --reuse
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
RENDERER = (
    REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
)
SCRATCH = Path(r"C:\tmp\xlsx_note_lines")

# Each arm: the face, the size, and how many lines the note holds. The text is
# all CJK so every line has the same ink top and bottom — a Latin line's ink
# starts at the cap height and would read as a different first baseline for no
# reason but the letters chosen.
# The fourth number is how tall the box is, in points. The corpus floor's note
# holds four lines of 14pt — 120 pixels of text — in a box 111 pixels high, so
# its text does not fit and every arm above did. Whether that is the fork is
# what the short boxes below ask.
ARMS = [
    ("ＭＳ Ｐゴシック", 9.0, 3, 120.0),
    ("ＭＳ Ｐゴシック", 9.0, 1, 120.0),
    ("ＭＳ Ｐゴシック", 11.0, 3, 120.0),
    ("ＭＳ Ｐゴシック", 14.0, 3, 120.0),
    ("ＭＳ Ｐゴシック", 18.0, 3, 120.0),
    ("ＭＳ ゴシック", 9.0, 3, 120.0),
    ("ＭＳ ゴシック", 14.0, 3, 120.0),
    ("游ゴシック", 9.0, 3, 120.0),
    ("游ゴシック", 14.0, 3, 120.0),
    ("メイリオ", 9.0, 3, 120.0),
    ("メイリオ", 14.0, 3, 120.0),
    # Boxes too small for what is in them.
    ("ＭＳ Ｐゴシック", 14.0, 4, 60.0),
    ("メイリオ", 14.0, 4, 60.0),
    ("メイリオ", 14.0, 4, 83.0),
    ("メイリオ", 14.0, 6, 83.0),
    ("游ゴシック", 14.0, 4, 60.0),
]
SAID = "国立国会図書館"
ROW_STEP = 12          # rows between one note and the next
WIDE = 200.0
TALL = 120.0           # the tallest box, which bounds the window read


def build(made: Path) -> None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:M400").Interior.Color = 0xFFFFFF
        for at, (face, size, lines, tall) in enumerate(ARMS):
            cell = sheet.Cells(2 + at * ROW_STEP, 2)
            note = cell.AddComment(chr(10).join([SAID] * lines))
            shape = note.Shape
            shape.Width = WIDE
            shape.Height = tall
            frame = shape.TextFrame
            frame.Characters().Font.Name = face
            frame.Characters().Font.Size = size
            note.Visible = True
        if made.exists():
            made.unlink()
        book.SaveAs(str(made), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def shoot(made: Path) -> bool:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(made))
    try:
        sheet = book.Worksheets(1)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(f"A1:M{4 + len(ARMS) * ROW_STEP}").CopyPicture(
                    Appearance=1, Format=2
                )
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(1.2)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def boxes(grey: np.ndarray) -> list[int]:
    """The top rule of every note in the picture.

    Found by the LENGTH of the rule rather than by looking where a note was
    asked for: Excel places a note beside its cell and moves it to keep it on
    the sheet, so where it ends up is not where the cell is. A cell holding a
    note also wears a small mark in its corner, and taking that for the box
    makes the two sides measure from different things.
    """
    rows = (grey < 128).sum(axis=1)
    long_ = []
    inside = False
    for y, held in enumerate(rows):
        if held > 100:
            if not inside:
                long_.append(y)
            inside = True
        else:
            inside = False
    # A note contributes two long rules, its top and its foot, exactly the
    # box's height apart. Pairing them is what tells a top from a foot; taking
    # every long rule for a top reads each note twice.
    held = set(long_)
    heights = {int(one[3] * 96 / 72) for one in ARMS}
    return [
        y
        for y in long_
        if any(y + tall + d in held for tall in heights for d in (-2, -1, 0, 1, 2))
    ]


def read(grey: np.ndarray, edge: int, tall: int) -> str:
    """Where each line of ink begins below a note's top rule."""
    band = grey[edge : edge + tall]
    rows = (band < 128).sum(axis=1)
    # More than the box's own side rules leave on a row, so an empty row
    # inside the note is not read as a line of text.
    lit = [i for i, v in enumerate(rows) if v > 8 and i > 2]
    if not lit:
        return "no text"
    runs, start, last = [], lit[0], lit[0]
    for i in lit[1:]:
        if i > last + 1:
            runs.append(start)
            start = i
        last = i
    runs.append(start)
    # The last run is the box's own bottom rule, not a line of text.
    if len(runs) > 1 and rows[runs[-1]] > 100:
        runs = runs[:-1]
    pitches = [runs[i + 1] - runs[i] for i in range(len(runs) - 1)]
    return (
        f"first +{runs[0]:<3} pitch "
        + (",".join(str(one) for one in pitches) if pitches else "-")
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "notes.xlsx"
    if not args.reuse:
        build(made)
        if not shoot(made):
            print("  Excel would not hand over a picture")
            return 1
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")).astype(int)
    drawing = dict(os.environ)
    drawing["OXI_XLSX_RANGE"] = f"1,1,{4 + len(ARMS) * ROW_STEP},13"
    subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=drawing,
    )
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")).astype(int)
    print(f"  Excel {truth.shape[1]}x{truth.shape[0]}, Oxi {mine.shape[1]}x{mine.shape[0]}")
    theirs, ours = boxes(truth), boxes(mine)
    print(f"  Excel drew {len(theirs)} box(es), Oxi {len(ours)}")
    tall = int(TALL * 96 / 72)
    print(f"  {'face':<16}{'pt':>5}{'n':>3}{'tall':>7}  {'box':>10}   {'Excel':<26}Oxi")
    agree = 0
    for at, (face, size, lines, tall) in enumerate(ARMS):
        if at >= len(theirs) or at >= len(ours):
            break
        room = int(tall * 96 / 72)
        one, two = read(truth, theirs[at], room), read(mine, ours[at], room)
        agree += one == two
        print(f"  {face:<16}{size:>5}{lines:>3}{tall:>7}  {theirs[at]:>5}/{ours[at]:<4}"
              f"   {one:<26}{two}{'' if one == two else '  <<'}")
    print(f"  {agree} of {min(len(ARMS), len(theirs), len(ours))} arms agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
