# -*- coding: utf-8 -*-
"""The phase rule, declared first and then measured on new arms.

From `_xlsx_shape_phase.py`, over ten faces and four sizes:

    A shape steps by round(design) a character and lets the run get AHEAD of
    its exact place. When the lead would pass a cap it takes one pixel back.
    The cap is 1.0 where the device's own advance is wider than round(design)
    — ＭＳ 明朝 and ＭＳ ゴシック, whose hinted advance is round(design)+1 for
    every character — and about 3.5 everywhere else.

So the first give-back lands at

    N = min{ k : k*(round(design) - design) > cap }

This declares N for arms the first sweep never saw, and then reads Excel's
picture to see whether N is where it was said to be. Both readings of the cap
are printed — 3.5 allowed, and 3.5 not allowed — because the first sweep's two
HGP arms disagreed by one about the boundary.

    python tools\\metrics\\_xlsx_shape_phase2.py
"""

from __future__ import annotations

import ctypes
import math
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
SCRATCH = Path(r"C:\tmp\xlsx_shape_phase2")
GDI = ctypes.windll.gdi32
USER = ctypes.windll.user32
LETTER = "日"          # one full-width ideograph that rasterises as ONE blob
                      # (相 is two — the radical and the body — and the reading
                      # of "where a glyph starts" then belongs to a half-glyph)
COUNT = 44
ARMS = [("AR P丸ゴシック体E", 12.0), ("AR P丸ゴシック体E", 14.0),
        ("ＤＦ特太ゴシック体", 14.0), ("メイリオ", 9.0), ("メイリオ", 11.0),
        ("メイリオ", 14.0), ("ＭＳ 明朝", 10.0), ("ＭＳ 明朝", 12.0), ("ＭＳ ゴシック", 16.0),
        ("メイリオ", 10.0), ("メイリオ", 12.0), ("メイリオ", 16.0), ("メイリオ", 18.0),
        ("游ゴシック", 12.0), ("游ゴシック", 16.0), ("游明朝", 12.0),
        ("ＭＳ Ｐゴシック", 12.0), ("ＭＳ Ｐゴシック", 16.0), ("ＭＳ Ｐ明朝", 12.0),
        ("Meiryo UI", 12.0), ("Meiryo UI", 16.0), ("BIZ UDPゴシック", 12.0),
        ("BIZ UDゴシック", 14.0), ("Yu Gothic UI", 14.0), ("HGP創英角ﾎﾟｯﾌﾟ体", 12.0),
        ("HGS明朝E", 14.0)]
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


def advance(face: str, height: int) -> int:
    dc = USER.GetDC(0)
    try:
        lf = LOGFONT()
        lf.lfHeight = height
        lf.lfCharSet = 128
        lf.lfFaceName = face[:31]
        font = GDI.CreateFontIndirectW(ctypes.byref(lf))
        old = GDI.SelectObject(dc, font)
        held = SIZE()
        GDI.GetTextExtentPoint32W(dc, LETTER, 1, ctypes.byref(held))
        GDI.SelectObject(dc, old)
        GDI.DeleteObject(font)
        return held.cx
    finally:
        USER.ReleaseDC(0, dc)


def said(face: str, points: float) -> tuple[float, int, float, int, int]:
    """design, round(design), cap, N if 3.5 is allowed, N if it is not."""
    em = points * 96 / 72
    design = advance(face, -2048) / 2048 * em
    whole = advance(face, -round(em))
    step = round(design)
    cap = 1.0 if whole > step else 3.5
    lead = step - design
    if lead <= 1e-6:
        return design, step, cap, 0, 0
    # N = the first k whose lead passes the cap.
    inclusive = math.floor(cap / lead) + 1          # cap itself allowed
    exclusive = math.ceil(cap / lead)               # cap itself not allowed
    return design, step, cap, inclusive, max(exclusive, 1)


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
        for face, points in ARMS:
            shape = sheet.Shapes.AddShape(1, 20.0, at, 1500.0, HIGH - 4)
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
    print("  face                size  design  step  cap   N said (3.5 in / out)   N read   verdict")
    tally = {"3.5 allowed": 0, "3.5 refused": 0, "neither": 0, "no give-back": 0}
    for face, points, top in placed:
        design, step, cap, inclusive, exclusive = said(face, points)
        seen = positions(truth, top)
        if len(seen) != COUNT:
            print(f"  {face:<18}{points:>5.0f}   {len(seen)} blobs for {COUNT} glyphs — not read")
            continue
        rel = [x - seen[0] for x in seen]
        fix = next((k for k in range(1, len(rel)) if rel[k] - rel[k - 1] != step), None)
        if fix is None:
            verdict = "no give-back"
        elif fix == inclusive == exclusive:
            verdict = "both agree"
        elif fix == inclusive:
            verdict = "3.5 allowed"
        elif fix == exclusive:
            verdict = "3.5 refused"
        else:
            verdict = "NEITHER"
        tally[verdict if verdict in tally else "neither"] = \
            tally.get(verdict if verdict in tally else "neither", 0) + 1
        print(f"  {face:<18}{points:>5.0f} {design:>7.3f} {step:>5} {cap:>4.1f}"
              f"   {inclusive:>6} / {exclusive:<6}      {str(fix):>5}   {verdict}")
    print(f"  {tally}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
