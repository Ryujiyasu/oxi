# -*- coding: utf-8 -*-
"""How far does a shape step per character, measured over a hundred of them?

`_xlsx_shape_advance2.py` asked the same question over 3 to 17 characters,
where the two models — a cell's whole-pixel advances, a shape's design-unit
accumulation — are five pixels apart, and Excel's answer landed BETWEEN them
(17 characters: 207 read, 210 cell, 205 shape). At ±1 pixel of edge that could
not be split.

So stretch the lever. At a hundred characters the models are thirty pixels
apart, and the per-character step falls out of the difference of two lengths
to a hundredth of a pixel:

    step = (reach(100) - reach(10)) / 90

Both lengths end on the same character, so the last glyph's own ink cancels.
One character repeated, so no pair kerning enters.

    python tools\\metrics\\_xlsx_shape_pitch.py
"""

from __future__ import annotations

import ctypes
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_shape_step")
LETTER = "あ"
END = "ぬ"
LENGTHS = (10, 30, 60, 100)
ARMS = [("メイリオ", 14.0), ("メイリオ", 11.0), ("ＭＳ Ｐゴシック", 14.0),
        ("ＭＳ ゴシック", 14.0), ("游ゴシック", 14.0), ("ＭＳ 明朝", 14.0)]


class LOGFONT(ctypes.Structure):
    _fields_ = [("lfHeight", ctypes.c_long), ("lfWidth", ctypes.c_long),
                ("lfEscapement", ctypes.c_long), ("lfOrientation", ctypes.c_long),
                ("lfWeight", ctypes.c_long), ("lfItalic", ctypes.c_byte),
                ("lfUnderline", ctypes.c_byte), ("lfStrikeOut", ctypes.c_byte),
                ("lfCharSet", ctypes.c_byte), ("lfOutPrecision", ctypes.c_byte),
                ("lfClipPrecision", ctypes.c_byte), ("lfQuality", ctypes.c_byte),
                ("lfPitchAndFamily", ctypes.c_byte), ("lfFaceName", ctypes.c_wchar * 32)]


class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


def steps(face: str, points: float) -> tuple[float, float]:
    """What the two models step per character: whole-pixel, and design."""
    gdi = ctypes.windll.gdi32
    user = ctypes.windll.user32
    dc = user.GetDC(0)
    try:
        def width(height: int) -> int:
            lf = LOGFONT()
            lf.lfHeight = height
            lf.lfCharSet = 128
            lf.lfFaceName = face[:31]
            font = gdi.CreateFontIndirectW(ctypes.byref(lf))
            old = gdi.SelectObject(dc, font)
            measured = SIZE()
            gdi.GetTextExtentPoint32W(dc, LETTER, 1, ctypes.byref(measured))
            gdi.SelectObject(dc, old)
            gdi.DeleteObject(font)
            return measured.cx

        em = points * 96 / 72
        return float(width(-round(em))), width(-2048) / 2048 * em
    finally:
        user.ReleaseDC(0, dc)


def picture(sheet):
    for _ in range(8):
        try:
            sheet.Activate()
            sheet.Range("A1:CZ30").CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.5)
        held = ImageGrab.grabclipboard()
        if held is not None:
            return held
    return None


def reach_of(sheet, words: str, face: str, points: float) -> int | None:
    shape = sheet.Shapes.AddShape(1, 40.0, 60.0, 2400.0, 44.0)   # 1 = rect
    try:
        frame = shape.TextFrame2
        frame.WordWrap = False
        frame.AutoSize = 0
        frame.VerticalAnchor = 1
        frame.TextRange.Text = words
        frame.TextRange.Font.Size = points
        # `Font.Name` sets only the LATIN face; kana keep the East Asian one.
        frame.TextRange.Font.Name = face
        try:
            frame.TextRange.Font.NameFarEast = face
        except Exception:
            pass
        frame.TextRange.Font.Bold = False
        frame.TextRange.Font.Fill.ForeColor.RGB = 0
        shape.Fill.Visible = False
        shape.Line.Visible = False
        held = picture(sheet)
        if held is None:
            return None
        top = round(shape.Top * 96 / 72)
        left = round(shape.Left * 96 / 72)
        grey = np.asarray(held.convert("L"))[
            top:top + round(shape.Height * 96 / 72), left:left + 2600
        ]
        ink = grey < 120
        if not ink.any():
            return None
        cols = np.where(ink.any(axis=0))[0]
        return int(cols.max() - cols.min())
    finally:
        shape.Delete()


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:CZ30").Interior.Color = 0xFFFFFF
        print(f"  one shape, {LETTER!r} repeated, every length ended by {END!r}")
        print("  face          size   read step   whole-pixel   design    verdict")
        for face, points in ARMS:
            seen: dict[int, int] = {}
            for count in LENGTHS:
                found = reach_of(sheet, LETTER * count + END, face, points)
                if found is None:
                    break
                seen[count] = found
            if len(seen) < len(LENGTHS):
                print(f"  {face} {points} — Excel drew nothing")
                continue
            short, long = LENGTHS[0], LENGTHS[-1]
            step = (seen[long] - seen[short]) / (long - short)
            whole, design = steps(face, points)
            which = ("whole-pixel" if abs(step - whole) < abs(step - design)
                     else "design" if abs(step - design) < abs(step - whole) else "tie")
            print(f"  {face:<14}{points:>4.0f}  {step:>9.3f}   {whole:>11.3f}"
                  f"   {design:>7.3f}   {which}")
            told = "     lengths: " + "  ".join(
                f"{count}->{seen[count]}" for count in LENGTHS)
            print(told)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
