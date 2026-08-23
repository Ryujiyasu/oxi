"""How does a cell's NOTE step from one character to the next?

`b6a3a84180c9_002` — the lowest-scoring workbook there is — draws its notes
with the right words, on the right lines, at the right pitch, and yet the text
creeps left the further along a line it goes: against Excel's own picture a
glyph run that starts together is three pixels short by the twentieth
character. Most advances agree and a few differ by exactly one, which is the
signature of rounding each ADVANCE rather than each POSITION.

A cell rounds advances (`advances`); a shape accumulates in the font's own
design units and rounds the position (`shape_run`). A note is drawn by the
cell engine. This asks Excel which of the two a note actually does.

Two traps this holds fixed, both of which produced confident nonsense first:
  * Excel puts the author's name at the head of a new note, bold and on its
    own line, so the first band of ink is not the text.
  * Glyph runs do not map one to one onto characters — neighbours touch, a
    letter breaks in two — so per-glyph positions cannot be read off. What is
    read instead is how far the line REACHES, which cannot merge.
An absolute reach is still unusable, because it carries the ink width of the
last glyph, which neither model predicts. So every length ends on the SAME
character and the reaches of two lengths are subtracted: the terminator's ink
cancels and what is left is purely where the two models put it.

Run: python tools/metrics/_xlsx_note_advance.py
"""

from __future__ import annotations

import ctypes
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_note_advance")
# Kana of one em each with narrow Latin letters mixed in, so whole-pixel
# advances and design-unit ones part company.
WORDS = "いAろBはCにDほEへFとGちHりI"
END = "ぬ"
LENGTHS = (3, 8, 13, 17)
ARMS = [("ＭＳ Ｐゴシック", 11.0), ("ＭＳ Ｐゴシック", 14.0), ("メイリオ", 11.0),
        ("メイリオ", 14.0), ("游ゴシック", 11.0), ("游ゴシック", 14.0)]


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


def where_last(face: str, points: float, text: str) -> tuple[int, int]:
    """Where the last character starts under each model, from the line's start."""
    gdi = ctypes.windll.gdi32
    user = ctypes.windll.user32
    dc = user.GetDC(0)
    try:
        def widths(height: int) -> list[int]:
            lf = LOGFONT()
            lf.lfHeight = height
            lf.lfCharSet = 128
            lf.lfFaceName = face[:31]
            font = gdi.CreateFontIndirectW(ctypes.byref(lf))
            old = gdi.SelectObject(dc, font)
            out = []
            for letter in text:
                measured = SIZE()
                gdi.GetTextExtentPoint32W(dc, letter, 1, ctypes.byref(measured))
                out.append(measured.cx)
            gdi.SelectObject(dc, old)
            gdi.DeleteObject(font)
            return out

        em = points * 96 / 72
        whole = widths(-round(em))
        design = [w / 2048 for w in widths(-2048)]
        # A cell: each advance a whole pixel, added up.
        cell = sum(whole[:-1])
        # A shape: the exact em accumulated, the POSITION rounded.
        at = 0.0
        for share in design[:-1]:
            at += share * em
        return cell, round(at)
    finally:
        user.ReleaseDC(0, dc)


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


def reach_of(sheet, target, words: str, face: str, points: float) -> int | None:
    """How far the note's line of ink reaches, in pixels."""
    if target.Comment is not None:
        target.Comment.Delete()
    note = target.AddComment(words)
    try:
        note.Visible = True
        frame = note.Shape.TextFrame
        frame.Characters().Text = words
        frame.Characters().Font.Name = face
        frame.Characters().Font.Size = points
        frame.Characters().Font.Bold = False
        note.Shape.Width = 900.0
        note.Shape.Height = 60.0
        note.Shape.Fill.ForeColor.RGB = 0xFFFFFF
        note.Shape.Line.Visible = False
        held = picture(sheet)
        if held is None:
            return None
        box = note.Shape
        top = round(box.Top * 96 / 72)
        left = round(box.Left * 96 / 72) + 4
        right = left + round(box.Width * 96 / 72) - 8
        grey = np.asarray(held.convert("L"))[top:top + round(box.Height * 96 / 72), left:right]
        ink = grey < 120
        if not ink.any():
            return None
        cols = np.where(ink.any(axis=0))[0]
        return int(cols.max() - cols.min())
    finally:
        if target.Comment is not None:
            target.Comment.Delete()


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
        print(f"note text {WORDS!r}, every length terminated by {END!r}")
        tally = {"CELL": 0, "SHAPE": 0, "tie": 0}
        for face, points in ARMS:
            seen: dict[int, int] = {}
            said: dict[int, tuple[int, int]] = {}
            for count in LENGTHS:
                words = WORDS[:count] + END
                said[count] = where_last(face, points, words)
                found = reach_of(sheet, target, words, face, points)
                if found is None:
                    print(f"  {face} {points:.0f}pt — Excel drew nothing at {count}")
                    break
                seen[count] = found
            if len(seen) < len(LENGTHS):
                continue
            base = LENGTHS[0]
            print(f"  {face} {points:.0f}pt")
            for count in LENGTHS[1:]:
                grew = seen[count] - seen[base]
                by_cell = said[count][0] - said[base][0]
                by_shape = said[count][1] - said[base][1]
                which = (
                    "CELL" if abs(grew - by_cell) < abs(grew - by_shape)
                    else "SHAPE" if abs(grew - by_shape) < abs(grew - by_cell)
                    else "tie"
                )
                tally[which] += 1
                print(f"    {base}->{count} chars: grew {grew:4d}"
                      f"   cell {by_cell:4d} ({grew - by_cell:+d})"
                      f"   shape {by_shape:4d} ({grew - by_shape:+d})   {which}")
        print(f"  verdict across arms: {tally}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
