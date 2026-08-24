# -*- coding: utf-8 -*-
"""How far does Excel let a shape's glyphs run ahead of their exact places?

`_xlsx_shape_step.py` settled the step: a shape advances by the DESIGN advance,
18.667 pixels a full-width character at 14 point. But the glyphs are put down on
WHOLE pixels, and Excel does not round each position the way we do. It steps by
the whole-pixel `round(design)` and lets the run get ahead of its exact place,
taking a pixel back only when the lead would pass a limit:

    excess(k) = drawn(k) - k*design    grows by round(design)-design a glyph,
                                       and is capped.

    ＭＳ明朝 14pt   cap 1.00   (0, .33, .67, 1.0, .33, .67, 1.0, ...)
    游ゴシック 14pt  cap 3.33   (0, .33, ... 3.33, 2.67, 3.0, 3.33, 2.67, ...)

Two faces with the SAME design advance and different caps, so the cap is not a
function of the advance. This sweeps faces and sizes to find what it IS.

    python tools\\metrics\\_xlsx_shape_phase.py
"""

from __future__ import annotations

import ctypes
import sys
import time
from ctypes import wintypes
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_shape_phase")
GDI = ctypes.windll.gdi32
USER = ctypes.windll.user32
LETTER = "あ"
COUNT = 40
FACES = ["ＭＳ 明朝", "ＭＳ ゴシック", "ＭＳ Ｐ明朝", "ＭＳ Ｐゴシック", "メイリオ",
         "Meiryo UI", "游ゴシック", "游明朝", "ＭＳ ＵＩ Ｇｏｔｈｉｃ", "HGP創英角ﾎﾟｯﾌﾟ体"]
SIZES = [9.0, 11.0, 14.0, 20.0]
TOP, HIGH = 30.0, 34.0


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


class TEXTMETRICW(ctypes.Structure):
    _fields_ = [("tmHeight", wintypes.LONG), ("tmAscent", wintypes.LONG),
                ("tmDescent", wintypes.LONG), ("tmInternalLeading", wintypes.LONG),
                ("tmExternalLeading", wintypes.LONG), ("tmAveCharWidth", wintypes.LONG),
                ("tmMaxCharWidth", wintypes.LONG), ("tmWeight", wintypes.LONG),
                ("tmOverhang", wintypes.LONG), ("tmDigitizedAspectX", wintypes.LONG),
                ("tmDigitizedAspectY", wintypes.LONG), ("tmFirstChar", wintypes.WCHAR),
                ("tmLastChar", wintypes.WCHAR), ("tmDefaultChar", wintypes.WCHAR),
                ("tmBreakChar", wintypes.WCHAR), ("tmItalic", wintypes.BYTE),
                ("tmUnderlined", wintypes.BYTE), ("tmStruckOut", wintypes.BYTE),
                ("tmPitchAndFamily", wintypes.BYTE), ("tmCharSet", wintypes.BYTE)]


def font_facts(face: str, points: float) -> dict[str, float]:
    """What the device says about the face at this size, and in design units."""
    dc = USER.GetDC(0)
    try:
        def measured(height: int) -> tuple[int, TEXTMETRICW]:
            lf = LOGFONT()
            lf.lfHeight = height
            lf.lfCharSet = 128
            lf.lfFaceName = face[:31]
            font = GDI.CreateFontIndirectW(ctypes.byref(lf))
            old = GDI.SelectObject(dc, font)
            width = SIZE()
            GDI.GetTextExtentPoint32W(dc, LETTER, 1, ctypes.byref(width))
            told = TEXTMETRICW()
            GDI.GetTextMetricsW(dc, ctypes.byref(told))
            GDI.SelectObject(dc, old)
            GDI.DeleteObject(font)
            return width.cx, told
        em = points * 96 / 72
        whole, told = measured(-round(em))
        design_units, _ = measured(-2048)
        return {"em": em, "whole": float(whole), "design": design_units / 2048 * em,
                "share": design_units / 2048, "ascent": told.tmAscent,
                "descent": told.tmDescent, "overhang": told.tmOverhang,
                "average": told.tmAveCharWidth, "widest": told.tmMaxCharWidth}
    finally:
        USER.ReleaseDC(0, dc)


def build() -> list[tuple[str, float, float]]:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    placed = []
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:BZ200").Interior.Color = 0xFFFFFF
        at = TOP
        for face in FACES:
            for points in SIZES:
                shape = sheet.Shapes.AddShape(1, 20.0, at, 1400.0, HIGH - 4)
                frame = shape.TextFrame2
                frame.WordWrap = False
                frame.AutoSize = 0
                frame.VerticalAnchor = 1
                frame.TextRange.Text = LETTER * COUNT
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
                placed.append((face, points, shape.Top))
                at += HIGH
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:BZ200").CopyPicture(Appearance=1, Format=2)
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
    return placed


def positions(picture: np.ndarray, top: float) -> list[int]:
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
    placed = build()
    if not placed:
        print("  Excel would not hand over a picture")
        return 1
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L"))
    print("  face                size   design  whole  step   first fix   cap"
          "    asc desc over  avg widest")
    for face, points, top in placed:
        seen = positions(truth, top)
        facts = font_facts(face, points)
        if len(seen) < 12:
            print(f"  {face:<18}{points:>5.0f}   only {len(seen)} glyphs read")
            continue
        rel = [x - seen[0] for x in seen]
        design = facts["design"]
        step = round(design)
        # Where the run first gives a pixel back.
        fix = next((k for k in range(1, len(rel)) if rel[k] - rel[k - 1] != step), None)
        cap = None if fix is None else round(rel[fix - 1] - (fix - 1) * design, 2)
        print(f"  {face:<18}{points:>5.0f}  {design:>7.3f} {facts['whole']:>6.0f}"
              f" {step:>5}   {str(fix):>9}  {str(cap):>5}"
              f"  {facts['ascent']:>4.0f} {facts['descent']:>4.0f} {facts['overhang']:>4.0f}"
              f" {facts['average']:>4.0f} {facts['widest']:>6.0f}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
