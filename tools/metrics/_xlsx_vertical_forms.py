# -*- coding: utf-8 -*-
"""Which characters does Excel turn on their side in a stacked cell, and how?

A stacked cell (`textRotation="255"`) stands its characters one above the
next, but not all of them upright: the long vowel mark, the brackets and the
dashes are drawn in their vertical forms. Every katakana word in a Japanese
form's column headings carries one, so drawing them flat is visible on every
one of the 36 workbooks that stack text.

This asks Excel for one character per cell and reads back the shape it drew,
against the same character drawn upright and drawn through the vertical
("@") face by the device.

    python tools\\metrics\\_xlsx_vertical_forms.py
"""
import argparse
import ctypes
import subprocess
import sys
from ctypes import wintypes
from pathlib import Path

import numpy as np
from PIL import Image

GDI = ctypes.windll.gdi32
USER = ctypes.windll.user32
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"
SCRATCH = Path(r"C:\tmp\xlsx_vertical")
BOOK = SCRATCH / "vertical.xlsx"
TRUTH = SCRATCH / "vertical.excel.png"

FACE = "ＭＳ ゴシック"
POINTS = 11.0
ROW_PX = 24
LETTERS = "アーｰ〜～（）「」【】ー・、。ヴァ゛＝ＡＢ１２"


def build():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["A"].width = 4
    for row, letter in enumerate(LETTERS, start=1):
        cell = sheet.cell(row=row, column=1, value=letter)
        cell.font = Font(name=FACE, size=POINTS)
        cell.alignment = Alignment(textRotation=255, vertical="center",
                                   horizontal="center")
        sheet.row_dimensions[row].height = ROW_PX * 0.75
    book.save(BOOK)
    return list(LETTERS)


def shoot():
    listing = SCRATCH / "_batch.txt"
    listing.write_text(f"{BOOK.resolve()}\t{TRUTH.resolve()}", encoding="utf-8")
    TRUTH.unlink(missing_ok=True)
    subprocess.run(["powershell", "-NoProfile", "-File", str(SHOOTER),
                    "-ListFile", str(listing.resolve())],
                   capture_output=True, text=True, encoding="utf-8",
                   errors="replace", timeout=300)
    listing.unlink(missing_ok=True)


def drawn_by_device(letter, face, points, upright=True):
    """The ink box of one character as the device draws it, upright or through
    the vertical face, as (width, height)."""
    pixels = round(points * 96.0 / 72.0)
    width = height = 80
    screen = USER.GetDC(None)
    dc = GDI.CreateCompatibleDC(screen)
    bitmap = GDI.CreateCompatibleBitmap(screen, width, height)
    GDI.SelectObject(dc, bitmap)
    GDI.PatBlt(dc, 0, 0, width, height, 0x00F00062)
    name = face if upright else "@" + face
    escapement = 0 if upright else 2700
    font = GDI.CreateFontW(-pixels, 0, escapement, escapement, 400, 0, 0, 0,
                           1, 0, 0, 5, 0, name)
    old = GDI.SelectObject(dc, font)
    GDI.SetBkMode(dc, 1)
    GDI.SetTextAlign(dc, 0)                 # TA_TOP | TA_LEFT
    GDI.TextOutW(dc, 30, 20, letter, len(letter))

    class HEADER(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                    ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                    ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                    ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                    ("biClrImportant", wintypes.DWORD)]

    header = HEADER()
    header.biSize = ctypes.sizeof(HEADER)
    header.biWidth, header.biHeight = width, -height
    header.biPlanes, header.biBitCount = 1, 32
    buffer = (ctypes.c_ubyte * (width * height * 4))()
    GDI.GetDIBits(dc, bitmap, 0, height, buffer, ctypes.byref(header), 0)
    GDI.SelectObject(dc, old)
    GDI.DeleteObject(font)
    GDI.DeleteObject(bitmap)
    GDI.DeleteDC(dc)
    USER.ReleaseDC(None, screen)

    grid = np.frombuffer(buffer, dtype=np.uint8).reshape(height, width, 4)[:, :, 0]
    lit = grid < 128
    columns = np.flatnonzero(lit.any(axis=0))
    rows = np.flatnonzero(lit.any(axis=1))
    if columns.size == 0:
        return (0, 0)
    return (int(columns[-1] - columns[0] + 1), int(rows[-1] - rows[0] + 1))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()

    letters = build()
    if not args.reuse:
        shoot()
    truth = np.asarray(Image.open(TRUTH).convert("L"))
    print(f"{'letter':>8}{'Excel w×h':>12}{'upright':>10}{'vertical face':>16}"
          f"  what Excel drew")
    for index, letter in enumerate(letters):
        band = truth[index * ROW_PX:(index + 1) * ROW_PX]
        lit = band < 128
        columns = np.flatnonzero(lit.any(axis=0))
        rows = np.flatnonzero(lit.any(axis=1))
        if columns.size == 0:
            print(f"{letter:>8}{'(no ink)':>12}")
            continue
        shape = (int(columns[-1] - columns[0] + 1), int(rows[-1] - rows[0] + 1))
        flat = drawn_by_device(letter, FACE, POINTS, upright=True)
        turned = drawn_by_device(letter, FACE, POINTS, upright=False)
        which = "upright" if shape == flat else ("turned" if shape == turned else "neither")
        print(f"{letter:>8}{f'{shape[0]}×{shape[1]}':>12}{f'{flat[0]}×{flat[1]}':>10}"
              f"{f'{turned[0]}×{turned[1]}':>16}  {which}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
