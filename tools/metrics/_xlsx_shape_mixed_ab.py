# -*- coding: utf-8 -*-
"""Does the phase rule hold for a line that is NOT all full-width?

The rule was read off runs of one repeated ideograph, where every advance is
the same. A real note is not that: `sanko_tool` and `001` mix Latin letters,
half-width digits and marks, whose design advances are fractions of an em and
whose per-character rounding is where the two models part company.

One shape a string, Excel's picture and ours, glyph start against glyph start.

    python tools\\metrics\\_xlsx_shape_mixed_ab.py
"""

from __future__ import annotations

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
SCRATCH = Path(r"C:\tmp\xlsx_shape_mixed_ab")
TOP, HIGH = 30.0, 34.0
ARMS = [
    ("メイリオ", 12.0, "The quick brown fox jumps over the lazy dog again and again"),
    ("メイリオ", 12.0, "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほ"),
    ("メイリオ", 12.0, "A1あB2いC3うD4えE5おF6かG7きH8くI9けJ0こKLMNOP"),
    ("メイリオ", 11.0, "The quick brown fox jumps over the lazy dog again and again"),
    ("メイリオ", 14.0, "A1あB2いC3うD4えE5おF6かG7きH8くI9けJ0こKLMNOP"),
    ("AR P丸ゴシック体E", 12.0, "A1あB2いC3うD4えE5おF6かG7きH8くI9けJ0こKLMNOP"),
    ("AR P丸ゴシック体E", 12.0, "参考ツールの使い方について（１）ここに説明が入ります。"),
    ("ＭＳ Ｐゴシック", 12.0, "A1あB2いC3うD4えE5おF6かG7きH8くI9けJ0こKLMNOP"),
    ("ＭＳ 明朝", 12.0, "A1あB2いC3うD4えE5おF6かG7きH8くI9けJ0こKLMNOP"),
    ("ＭＳ 明朝", 14.0, "A1あB2いC3うD4えE5おF6かG7きH8くI9けJ0こKLMNOP"),
]


def build(made: Path) -> list[float]:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    tops = []
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:BZ100").Interior.Color = 0xFFFFFF
        at = TOP
        for face, points, words in ARMS:
            shape = sheet.Shapes.AddShape(1, 20.0, at, 1500.0, HIGH - 4)
            frame = shape.TextFrame2
            frame.WordWrap = False
            frame.AutoSize = 0
            frame.VerticalAnchor = 1
            frame.TextRange.Text = words
            frame.TextRange.Font.Size = points
            frame.TextRange.Font.Name = face
            try:
                frame.TextRange.Font.NameFarEast = face
            except Exception:
                pass
            frame.TextRange.Font.Bold = False
            frame.TextRange.Font.Fill.ForeColor.RGB = 0
            shape.Fill.Visible = False
            shape.Line.Visible = False
            tops.append(shape.Top)
            at += HIGH
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:BZ100").CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.8)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                break
        else:
            return []
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return tops


def starts(picture: np.ndarray, top: float) -> list[int]:
    band = picture[round(top * 96 / 72):round((top + HIGH - 4) * 96 / 72)]
    lit = (band < 120).any(axis=0)
    out, start = [], None
    for at, held in enumerate(lit):
        if held and start is None:
            start = at
        elif not held and start is not None:
            out.append(start)
            start = None
    return out


def main() -> int:
    made = SCRATCH / "mixed.xlsx"
    tops = build(made)
    if not tops:
        print("  Excel would not hand over a picture")
        return 1
    ours = SCRATCH / "oxi.png"
    subprocess.run([str(RENDERER), str(made), str(ours), "96"], capture_output=True, check=False)
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L"))
    mine = np.asarray(Image.open(ours).convert("L"))
    print("  face            size  blobs   same   worst   text")
    for (face, points, words), top in zip(ARMS, tops):
        theirs, ours_at = starts(truth, top), starts(mine, top)
        count = min(len(theirs), len(ours_at))
        if count < 5:
            print(f"  {face:<15}{points:>4.0f}   nothing to read")
            continue
        moved = [(ours_at[i] - ours_at[0]) - (theirs[i] - theirs[0]) for i in range(count)]
        same = sum(1 for step in moved if step == 0)
        worst = max(moved, key=abs)
        print(f"  {face:<15}{points:>4.0f}  {len(theirs):>3}/{len(ours_at):<3} {same:>5}"
              f"  {worst:>+6}   {words[:22]}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
