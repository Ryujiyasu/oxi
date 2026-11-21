# -*- coding: utf-8 -*-
"""Probe: the room an English-mode Excel wraps a cell's text against.

For each (face, size, text) the column is swept a pixel at a time and the
narrowest column that keeps the text on one line is found (row height drops
to one line). Printed beside the text's GDI width at the cell's pixel em, so
the room Excel keeps (column - threshold width) can be read off.

    python tools/metrics/_xlsx_wrap_room_latn.py
"""
import ctypes, ctypes.wintypes as W, json
from pathlib import Path
import win32com.client

REPO = Path(__file__).resolve().parents[2]
import sys
ARMS = [("Calibri", 11, "n " + "n" * k) for k in (2, 6, 12, 20, 30, 45)] + [("Arial", 10, "n " + "n" * k) for k in (2, 12, 30, 45)] if "--sweep" in sys.argv else [("Calibri", 10, "Dayton 1200r24   DA53  THI"), ("Calibri", 10, "abcdefghij klm"), ("Calibri", 11, "abcdefghij klm"),
        ("Arial", 10, "abcdefghij klm"), ("Times New Roman", 12, "ab cd"), ("Times New Roman", 12, "abcdefghij klm"),
        ("Verdana", 9, "abcdefghij klm"), ("Calibri", 14, "abcdefghij klm")]

gdi = ctypes.windll.gdi32


class LOGFONTW(ctypes.Structure):
    _fields_ = [('lfHeight', W.LONG), ('lfWidth', W.LONG), ('lfEscapement', W.LONG), ('lfOrientation', W.LONG),
                ('lfWeight', W.LONG), ('lfItalic', W.BYTE), ('lfUnderline', W.BYTE), ('lfStrikeOut', W.BYTE),
                ('lfCharSet', W.BYTE), ('lfOutPrecision', W.BYTE), ('lfClipPrecision', W.BYTE), ('lfQuality', W.BYTE),
                ('lfPitchAndFamily', W.BYTE), ('lfFaceName', W.WCHAR * 32)]


def gdi_width(face, points, text):
    dc = gdi.CreateCompatibleDC(0)
    lf = LOGFONTW(); lf.lfHeight = -round(points * 96 / 72); lf.lfWeight = 400; lf.lfCharSet = 1; lf.lfFaceName = face
    f = gdi.CreateFontIndirectW(ctypes.byref(lf)); old = gdi.SelectObject(dc, f)
    sz = W.SIZE(); gdi.GetTextExtentPoint32W(dc, text, len(text), ctypes.byref(sz))
    gdi.SelectObject(dc, old); gdi.DeleteObject(f); gdi.DeleteDC(dc)
    return sz.cx


def main():
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    out = []
    try:
        wb = xl.Workbooks.Add(); sh = wb.Worksheets(1)
        for face, size, text in ARMS:
            c = sh.Range("A1"); c.Value = text; c.Font.Name = face; c.Font.Size = size; c.WrapText = True
            g = gdi_width(face, size, text)
            one = None
            for px in range(max(10, g - 4), g + int(g * 0.05) + 12):
                sh.Columns(1).ColumnWidth = 1  # reset
                # set width in points directly is not possible; find char width giving this px
                lo, hi = 0.0, 255.0
                for _ in range(40):
                    mid = (lo + hi) / 2
                    sh.Columns(1).ColumnWidth = mid
                    if round(sh.Columns(1).Width / 0.75) < px: lo = mid
                    else: hi = mid
                sh.Columns(1).ColumnWidth = hi
                col = round(sh.Columns(1).Width / 0.75)
                sh.Rows(1).AutoFit()
                h = round(sh.Rows(1).Height / 0.75)
                if one is None or h < one[1]:
                    pass
                out.append((face, size, text, g, col, h))
            sh.Range("A1").Clear()
        wb.Close(False)
    finally:
        xl.Quit()
    for face, size, text in ARMS:
        rows = [(col, h) for f, s, t, g, col, h in out if (f, s, t) == (face, size, text)]
        g = [g for f, s, t, g, col, h in out if (f, s, t) == (face, size, text)][0]
        hmin = min(h for _, h in rows)
        first = min(col for col, h in rows if h == hmin)
        print(f"{face:16s}{size:3} {text!r:30s} gdi {g:4d}  one line from col {first}px  -> room kept {first - g}")


if __name__ == "__main__":
    main()
