# -*- coding: utf-8 -*-
"""Where is Excel's pen when it puts a stacked character down?

Two faces draw a stacked cell: the upright one for what a vertical page leaves
standing, and the turned one — "@ＭＳ 明朝" a quarter turn — for the marks it
lays on their side. Each has its own idea of where the pen sits relative to the
ink, so Excel's own pen can be read back from the ink it leaves:

    pen = (ink Excel shows) - (offset GDI gives that face at that size)

Read against the cell's top-left, the em, and the face's line box, that says
what Excel's rule is — not a per-size correction table.

    python tools\\metrics\\_xlsx_stack_pen.py
"""

from __future__ import annotations

import ctypes
import re
import subprocess
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _xlsx_stack_face as FACE  # noqa: E402  (the DIB helpers live there)

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_stack_pen")
GDI = ctypes.windll.gdi32

FACE_NAME = "ＭＳ 明朝"
SIZES = [6.0, 8.0, 9.0, 11.0, 14.0, 20.0, 36.0]
UPRIGHT, TURNED = "相", "ー"
ROW_PT = 60.0
COLUMN = 9.0
TA_TOP, TA_BASELINE = 0, 24


def offset(face: str, px: int, letter: str, turned: bool, align: int) -> tuple[int, int, int, int] | None:
    """The ink's box relative to the pen, for one face at one pixel size."""
    pen = (40, 40)
    side = 120
    dc = GDI.CreateCompatibleDC(None)
    info = FACE.BITMAPINFO()
    head = info.bmiHeader
    head.biSize = ctypes.sizeof(FACE.BITMAPINFOHEADER)
    head.biWidth, head.biHeight = side, -side
    head.biPlanes, head.biBitCount, head.biCompression = 1, 32, 0
    bits = ctypes.c_void_p()
    bitmap = GDI.CreateDIBSection(dc, ctypes.byref(info), 0, ctypes.byref(bits), None, 0)
    GDI.SelectObject(dc, bitmap)
    room = (ctypes.c_uint32 * (side * side)).from_address(bits.value)
    for at in range(side * side):
        room[at] = 0x00FFFFFF
    turn = 2700 if turned else 0
    name = f"@{face}" if turned else face
    font = GDI.CreateFontW(-px, 0, turn, turn, 400, 0, 0, 0, 1, 0, 0, 5, 0, name)
    GDI.SelectObject(dc, font)
    GDI.SetTextColor(dc, 0x000000)
    GDI.SetBkMode(dc, 1)
    GDI.SetTextAlign(dc, align)
    held = ctypes.create_unicode_buffer(letter)
    GDI.TextOutW(dc, pen[0], pen[1], held, len(letter))
    seen = np.frombuffer(bytes(bytearray(room)), dtype=np.uint32).reshape(side, side)
    lit = (((seen & 0xFF) + ((seen >> 8) & 0xFF) + ((seen >> 16) & 0xFF)) // 3) < 140
    GDI.DeleteObject(font)
    GDI.DeleteObject(bitmap)
    GDI.DeleteDC(dc)
    rows, cols = np.where(lit)
    if len(rows) == 0:
        return None
    return (int(cols.min()) - pen[0], int(rows.min()) - pen[1],
            int(cols.max() - cols.min() + 1), int(rows.max() - rows.min() + 1))


def ascent(face: str, px: int, turned: bool) -> int:
    """How far the face's top sits above its baseline, as the device draws it."""
    top = offset(face, px, UPRIGHT if not turned else "一", turned, TA_TOP)
    base = offset(face, px, UPRIGHT if not turned else "一", turned, TA_BASELINE)
    if top is None or base is None:
        return 0
    return (base[0] - top[0]) if turned else (base[1] - top[1]) * -1


def shot() -> tuple[np.ndarray, int, dict[int, tuple[int, int]]] | None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "pen.xlsx"
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range(f"A1:E{len(SIZES) + 4}").Interior.Color = 0xFFFFFF
        for column in ("B", "C"):
            sheet.Columns(column).ColumnWidth = COLUMN
        for at, size in enumerate(SIZES, start=2):
            for column, letter in ((2, UPRIGHT), (3, TURNED)):
                cell = sheet.Cells(at, column)
                cell.Value = letter
                cell.Font.Name = FACE_NAME
                cell.Font.Size = size
                cell.Orientation = -4166          # xlVertical
                cell.VerticalAlignment = -4160    # xlTop
                cell.HorizontalAlignment = -4108  # xlCenter
            sheet.Rows(at).RowHeight = ROW_PT
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(SIZES), 3))
        book.SaveAs(str(made), FileFormat=51)
        for _ in range(12):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(0.6)
            grabbed = ImageGrab.grabclipboard()
            if grabbed is not None:
                break
        else:
            return None
        grabbed.save(SCRATCH / "excel.png")
        wide = round(sheet.Cells(2, 2).Width * 96 / 72)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    told = subprocess.run([str(RENDERER), str(made), str(SCRATCH / "x.png")],
                          env={"OXI_XLSX_DUMP_ROWS": "1", **dict(__import__("os").environ)},
                          capture_output=True, text=True, encoding="utf-8").stdout
    rows, at = {}, 0
    for line in told.splitlines():
        found = re.match(r"row (\d+) px (\d+)", line)
        if found:
            rows[int(found.group(1))] = (at, at + int(found.group(2)))
            at += int(found.group(2))
    truth = np.asarray(Image.open(SCRATCH / "excel.png").convert("L")) < 140
    return truth, wide, rows


def inked(block: np.ndarray) -> tuple[int, int] | None:
    rows, cols = np.where(block)
    if len(rows) == 0:
        return None
    return int(cols.min()), int(rows.min())


def main() -> int:
    got = shot()
    if got is None:
        print("  Excel would not hand over a picture")
        return 1
    truth, wide, ruled = got
    tall = round(ROW_PT * 96 / 72)
    print(f"  {FACE_NAME}, column {wide}px, row {tall}px, both cells centred, sat on the top")
    print("  size  em  box  asc_up asc_turn |  pen_up-top  box-em  |  pen_turn-top  |"
          "  pen_up-left  (w-em)/2  |  pen_turn-left  desc_turn")
    for at, size in enumerate(SIZES):
        px = round(size * 96 / 72)
        band = truth[at * tall:(at + 1) * tall]
        up = inked(band[:, :wide])
        turn = inked(band[:, wide:2 * wide])
        if up is None or turn is None:
            print(f"  {size:>5}  nothing to read")
            continue
        up_off = offset(FACE_NAME, px, UPRIGHT, False, TA_TOP)
        turn_off = offset(FACE_NAME, px, TURNED, True, TA_BASELINE)
        asc_up, asc_turn = ascent(FACE_NAME, px, False), ascent(FACE_NAME, px, True)
        pen_up = (up[0] - up_off[0], up[1] - up_off[1])
        pen_turn = (turn[0] - turn_off[0] + wide, turn[1] - turn_off[1])
        box = line_box_px(size)
        print(f"  {size:>5} {px:>3} {box:>4}  {asc_up:>6} {asc_turn:>8} |"
              f"  {pen_up[1]:>10} {box - px:>7}  |  {pen_turn[1]:>13}  |"
              f"  {pen_up[0]:>11} {round((wide - px) / 2):>9}  |"
              f"  {pen_turn[0] - wide:>14} {px - asc_turn:>10}")
    return 0


def line_box_px(size: float) -> int:
    """The row box the renderer's own measured table gives this face and size."""
    table = REPO / "tools" / "oxi-xlsx-renderer" / "src" / "row_defaults.rs"
    for line in table.read_text(encoding="utf-8").splitlines():
        found = re.search(r'\("([^"]+)",\s*(\d+),\s*(\d+),\s*(\d+)\)', line)
        if found and found.group(1) == FACE_NAME and int(found.group(2)) == round(size * 4):
            return int(found.group(3))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
