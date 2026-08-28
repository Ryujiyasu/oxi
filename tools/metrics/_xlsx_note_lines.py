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
# The top margin is NOT swept here. Arms added past the sixteenth fall off the
# bottom of the picture Excel hands over, and the reader then pairs them with
# earlier notes and reports four identical rows — which read as "Excel ignores
# the inset" and is not true. `_xlsx_note_002.py` asks that of the real note
# instead, and Excel moves its first line one pixel for one pixel of inset.
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


def notes(rgb: np.ndarray) -> list[tuple[int, int, int, int]]:
    """Every note in the picture, as the box its FILL covers.

    Found by the fill rather than by the rules around it. Pairing a top rule
    with a foot rule needs to know how tall the box is, which is one of the
    things being varied, and the pairing came apart as soon as two heights were
    in play — Excel's side found twenty-one boxes where ours found sixteen and
    the arms after the first few were reading different notes.

    A note's paper is the one colour on the sheet: 255,255,225, which Excel and
    the renderer both use for a note that names no fill of its own.
    """
    paper = (rgb[:, :, 0] == 255) & (rgb[:, :, 1] == 255) & (rgb[:, :, 2] == 225)
    wide = paper.sum(axis=1)
    bands, run = [], []
    for y, held in enumerate(wide):
        if held > 40:
            run.append(y)
        elif run:
            bands.append(run)
            run = []
    if run:
        bands.append(run)
    out = []
    for band in bands:
        columns = np.nonzero(paper[band[0] : band[-1] + 1].any(axis=0))[0]
        out.append((band[0], band[-1], int(columns[0]), int(columns[-1])))
    return out


def read(rgb: np.ndarray, note: tuple[int, int, int, int]) -> str:
    """Where each line of ink begins, counted from the top of the paper."""
    top, bottom, left, right = note
    ink = (rgb[top : bottom + 1, left : right + 1] < 128).all(axis=2)
    rows = ink.sum(axis=1)
    # More than the box's own side rules leave on a row, so an empty row inside
    # the note is not read as a line of text.
    lit = [i for i, v in enumerate(rows) if v > 8]
    if not lit:
        return "no text"
    runs, start, last = [], lit[0], lit[0]
    for i in lit[1:]:
        if i > last + 1:
            runs.append(start)
            start = i
        last = i
    runs.append(start)
    pitches = [runs[i + 1] - runs[i] for i in range(len(runs) - 1)]
    return (
        f"first +{runs[0]:<3} lines {len(runs)}  pitch "
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
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("RGB")).astype(int)
    drawing = dict(os.environ)
    drawing["OXI_XLSX_RANGE"] = f"1,1,{4 + len(ARMS) * ROW_STEP},13"
    subprocess.run(
        [str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=drawing,
    )
    mine = np.asarray(Image.open(SCRATCH / "oxi.png").convert("RGB")).astype(int)
    theirs, ours = notes(truth), notes(mine)
    print(f"  Excel {truth.shape[1]}x{truth.shape[0]}, Oxi {mine.shape[1]}x{mine.shape[0]}")
    print(f"  Excel drew {len(theirs)} note(s), Oxi {len(ours)}")
    if len(theirs) != len(ours):
        print("  the two sides do not hold the same notes; nothing to compare")
        return 1
    print(f"  {'face':<16}{'pt':>5}{'n':>3}{'tall':>7}  {'paper':>11}"
          f"  {'Excel':<30}Oxi")
    agree = 0
    for at, (face, size, lines, tall) in enumerate(ARMS):
        if at >= len(theirs):
            break
        one, two = read(truth, theirs[at]), read(mine, ours[at])
        agree += one == two
        print(f"  {face:<16}{size:>5}{lines:>3}{tall:>7}"
              f"  {theirs[at][0]:>5}/{ours[at][0]:<5}"
              f"  {one:<30}{two}{'' if one == two else '  <<'}")
    print(f"  {agree} of {min(len(ARMS), len(theirs))} arms agree")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
