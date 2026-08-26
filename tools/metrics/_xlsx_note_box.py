# -*- coding: utf-8 -*-
r"""How tall is the box Excel draws for a note?

The corpus floor, `b6a3a84180c9_002`, drops a whole line of a cell comment:
Excel draws the note's second line, we do not. The reason is the box — Excel's
is 79 pixels where ours is 77, and the line needs those two.

The VML states the size in points (`width:175.5pt;height:59pt`), so the
question is what Excel does with that number. 59pt is 78.67px at 96dpi, which
rounds to 79 and truncates to 78 — neither is 77, and the sheet holds notes of
two different heights, so the arithmetic cannot be settled by reading the file.

This asks instead: one note an arm, its height set through COM, and the box
read straight off Excel's own picture. The border is what is measured — a note
is filled, so its edge is where the fill stops.

    python tools\metrics\_xlsx_note_box.py
    python tools\metrics\_xlsx_note_box.py --reuse
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
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_note_box")
HEIGHTS = [30.0, 40.0, 45.75, 50.0, 57.75, 58.0, 58.5, 59.0, 59.25, 60.0, 72.0, 82.5]
WIDE = 120.0
GAP = 3          # rows between one note and the next


def build() -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:H400").Interior.Color = 0xFFFFFF
        sheet.Columns(1).ColumnWidth = 3
        sheet.Columns(2).ColumnWidth = 30
        at = 2
        for tall in HEIGHTS:
            cell = sheet.Cells(at, 2)
            cell.ClearComments()
            note = cell.AddComment("x")
            note.Shape.Width = WIDE
            note.Shape.Height = tall
            note.Shape.Top = sheet.Cells(at, 2).Top
            note.Shape.Left = sheet.Cells(at, 2).Left
            note.Visible = True
            at += GAP + int(tall / 15) + 2
        book.SaveAs(str(SCRATCH / "notes.xlsx"), FileFormat=51)
        rows = at + 4
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(rows, 8)).CopyPicture(
                    Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.9)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def boxes(picture: Image.Image) -> list[tuple[int, int]]:
    """Every note's top and bottom, found by its fill."""
    rgb = np.asarray(picture.convert("RGB")).astype(int)
    # Excel's note is #FFFFE1 unless told otherwise. The tolerance has to be
    # tight: white is only thirty away from it, so a loose one calls the whole
    # sheet a note.
    fill = (np.abs(rgb - np.array([255, 255, 225])).sum(axis=2) < 12)
    rows = np.where(fill.sum(axis=1) > 20)[0]
    out, start, last = [], None, None
    def shut(top: int, foot: int) -> None:
        cols = np.where(fill[top:foot + 1].any(axis=0))[0]
        out.append((int(top), int(foot), int(cols.min()), int(cols.max())))
    for y in rows:
        if start is None:
            start = last = y
            continue
        if y > last + 1:
            shut(start, last)
            start = y
        last = y
    if start is not None:
        shut(start, last)
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    if not args.reuse and not build():
        print("  Excel would not hand over a picture")
        return 1
    found = boxes(Image.open(SCRATCH / "excel.png"))
    subprocess.run(
        [str(RENDERER), str(SCRATCH / "notes.xlsx"), str(SCRATCH / "oxi.png"), "96"],
        capture_output=True, text=True, encoding="utf-8", env=dict(os.environ),
    )
    ours = boxes(Image.open(SCRATCH / "oxi.png")) if (SCRATCH / "oxi.png").exists() else []
    print(f"  {len(HEIGHTS)} note(s) asked for; Excel drew {len(found)}, we drew {len(ours)}")
    print(f"  {'asked pt':>9}{'96dpi px':>10}{'tall E':>7}{'O':>6}{'wide E':>9}{'O':>6}{'top E':>7}{'O':>5}{'left E':>7}{'O':>5}   what Excel's is")
    for at, (tall, held) in enumerate(zip(HEIGHTS, found)):
        drawn = held[1] - held[0] + 1
        across = held[3] - held[2] + 1
        mine = ours[at][1] - ours[at][0] + 1 if at < len(ours) else None
        mine_across = ours[at][3] - ours[at][2] + 1 if at < len(ours) else None
        exact = tall * 96 / 72
        # The fill stops one inside the border on each side, so the box the
        # renderer is asked for is two more than the fill.
        note = []
        for name, value in (("round", round(exact)), ("floor", int(exact)),
                            ("ceil", -(-exact // 1))):
            if drawn == value:
                note.append(name)
            if drawn == value + 2:
                note.append(f"{name}+2")
        top_e, top_o = held[0], ours[at][0] if at < len(ours) else None
        left_e, left_o = held[2], ours[at][2] if at < len(ours) else None
        print(f"  {tall:>9}{exact:>10.2f}{drawn:>7}{str(mine):>6}"
              f"{across:>9}{str(mine_across):>6}"
              f"{top_e:>7}{str(top_o):>5}{left_e:>7}{str(left_o):>5}   "
              f"{', '.join(note) or '—'}"
              f"{'' if (mine, mine_across, top_o, left_o) == (drawn, across, top_e, left_e) else '  <<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
