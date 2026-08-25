# -*- coding: utf-8 -*-
r"""Which gap gets the odd pixel when a distributed line is spread?

Reading this off a real workbook does not work: the pieces are different
glyphs, and where their INK starts inside their own advance differs, so the
gaps inferred from the picture carry each glyph's own side bearing. `fies_t2`
gives 39,39,40 on one line and 40,40,39 on another for that reason.

So the arms are identical glyphs, whose bearings cancel, in a column swept a
pixel at a time — the spare runs through every remainder against the number of
gaps, and where the odd pixel falls is the rounding. Three, four and five
pieces stand together, because a remainder of one and a remainder of two need
not fall the same way.

    python tools\metrics\_xlsx_distributed_round.py
    python tools\metrics\_xlsx_distributed_round.py --reuse
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
SCRATCH = Path(r"C:\tmp\xlsx_distributed_round")
FACE, POINTS = "ＭＳ ゴシック", 9.0
ROW_PT = 18.0
WORDS = {3: "日日日", 4: "日日日日", 5: "日日日日日"}
WIDTHS = [14.0 + step / 8.0 for step in range(16)]


def build(made: Path) -> bool:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:AZ12").Interior.Color = 0xFFFFFF
        # One arm a COLUMN, because a column can only hold one width.
        for at, width in enumerate(WIDTHS, start=2):
            sheet.Columns(at).ColumnWidth = width
            for row, count in enumerate(sorted(WORDS), start=2):
                cell = sheet.Cells(row, at)
                cell.NumberFormat = "@"
                cell.Value = WORDS[count]
                cell.Font.Name = FACE
                cell.Font.Size = POINTS
                cell.HorizontalAlignment = -4117        # xlDistributed
                cell.VerticalAlignment = -4108
        for row in range(2, 2 + len(WORDS)):
            sheet.Rows(row).RowHeight = ROW_PT
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range(sheet.Cells(2, 2),
                            sheet.Cells(1 + len(WORDS), 1 + len(WIDTHS))).CopyPicture(
                    Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.8)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def ours(made: Path):
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "oxi.png"), "96"],
                          env={"OXI_XLSX_DUMP_COLUMNS": "1", "OXI_XLSX_DUMP_ROWS": "1",
                               **os.environ},
                          capture_output=True, text=True, encoding="utf-8")
    columns, rows, at, down = {}, {}, 0, 0
    for line in (told.stdout or "").splitlines() + (told.stderr or "").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column":
            columns[int(parts[1])] = (at, at + int(parts[3]))
            at += int(parts[3])
        if len(parts) == 4 and parts[0] == "row":
            rows[int(parts[1])] = (down, down + int(parts[3]))
            down += int(parts[3])
    return np.asarray(Image.open(SCRATCH / "oxi.png").convert("L")) < 140, columns, rows


def starts(band: np.ndarray) -> list[int]:
    """Where each run of ink begins, across the band."""
    col = band.any(axis=0)
    out, at = [], False
    for i, v in enumerate(col):
        if v and not at:
            out.append(i)
        at = v
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    made = SCRATCH / "round.xlsx"
    if not args.reuse and not build(made):
        print("  Excel would not hand over a picture")
        return 1
    mine, columns, rows = ours(made)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    print(f"  {FACE} {POINTS}pt distributed, identical glyphs")
    print("  cell px  n |  gaps: Excel            Oxi")
    walked = 0
    for at in range(len(WIDTHS)):
        left, right = columns[at + 1]
        wide = right - left
        for step, count in enumerate(sorted(WORDS)):
            top, foot = rows[step + 2]
            tall = foot - top
            theirs = starts(truth[step * tall + 1:(step + 1) * tall - 1,
                                  walked + 1:walked + wide - 1])
            ours_at = starts(mine[top + 1:foot - 1, left + 1:right - 1])
            if len(theirs) != count or len(ours_at) != count:
                print(f"  {wide:>7} {count:>2} |  read {len(theirs)}/{len(ours_at)} runs, not {count}")
                continue
            gt = [theirs[i + 1] - theirs[i] for i in range(count - 1)]
            go = [ours_at[i + 1] - ours_at[i] for i in range(count - 1)]
            print(f"  {wide:>7} {count:>2} |  {str(gt):<22} {str(go):<22}"
                  f" {'' if gt == go else '<<'}")
        walked += wide
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
