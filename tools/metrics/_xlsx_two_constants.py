# -*- coding: utf-8 -*-
"""Are Excel's two unexplained per-font constants the same quantity?

One is the height a row takes beyond the font's own ascent and descent —
ＭＳ 明朝 +3, Century +2, 游ゴシック +5 (SX20). The other is the width a wrapped
line keeps beyond its text — 5px on small fonts, 7 and up on larger ones,
with faces at the same em disagreeing (SX17). Both are small, both grow with
the font, and neither follows any single TEXTMETRIC field.

Measured side by side on the same faces, with the full metrics beside them,
so a relation has somewhere to show itself.
"""
import ctypes
import ctypes.wintypes as w
import json
import sys

import win32com.client

gdi = ctypes.windll.gdi32
user = ctypes.windll.user32


class TEXTMETRICW(ctypes.Structure):
    _fields_ = [
        ("tmHeight", w.LONG), ("tmAscent", w.LONG), ("tmDescent", w.LONG),
        ("tmInternalLeading", w.LONG), ("tmExternalLeading", w.LONG),
        ("tmAveCharWidth", w.LONG), ("tmMaxCharWidth", w.LONG),
        ("tmWeight", w.LONG), ("tmOverhang", w.LONG),
        ("tmDigitizedAspectX", w.LONG), ("tmDigitizedAspectY", w.LONG),
        ("tmFirstChar", w.WCHAR), ("tmLastChar", w.WCHAR),
        ("tmDefaultChar", w.WCHAR), ("tmBreakChar", w.WCHAR),
        ("tmItalic", ctypes.c_byte), ("tmUnderlined", ctypes.c_byte),
        ("tmStruckOut", ctypes.c_byte), ("tmPitchAndFamily", ctypes.c_byte),
        ("tmCharSet", ctypes.c_byte),
    ]


class SIZE(ctypes.Structure):
    _fields_ = [("cx", w.LONG), ("cy", w.LONG)]


def font_facts(face, points, sample):
    hdc = user.GetDC(0)
    font = gdi.CreateFontW(-int(round(points * 96.0 / 72.0)), 0, 0, 0, 400,
                           0, 0, 0, 1, 0, 0, 0, 0, face)
    old = gdi.SelectObject(hdc, font)
    tm = TEXTMETRICW()
    gdi.GetTextMetricsW(hdc, ctypes.byref(tm))
    size = SIZE()
    body = ctypes.create_unicode_buffer(sample)
    gdi.GetTextExtentPoint32W(hdc, body, len(sample), ctypes.byref(size))
    got = ctypes.create_unicode_buffer(64)
    gdi.GetTextFaceW(hdc, 64, got)
    gdi.SelectObject(hdc, old)
    gdi.DeleteObject(font)
    user.ReleaseDC(0, hdc)
    return tm, size.cx, got.value


def narrowest_fit(ws):
    def height_at(chars):
        ws.Columns(1).ColumnWidth = max(chars, 0.2)
        ws.Rows(1).AutoFit()
        return ws.Rows(1).Height

    wide, narrow = 120.0, 1.0
    one_line = height_at(wide)
    for _ in range(26):
        middle = (narrow + wide) / 2
        if height_at(middle) > one_line + 0.01:
            narrow = middle
        else:
            wide = middle
    height_at(wide)
    return round(ws.Columns(1).Width / 0.75)


def main():
    faces = [
        ("ＭＳ 明朝", 9), ("ＭＳ 明朝", 11), ("ＭＳ 明朝", 12), ("ＭＳ 明朝", 14),
        ("ＭＳ Ｐゴシック", 11), ("ＭＳ Ｐゴシック", 14),
        ("游ゴシック", 9), ("游ゴシック", 11), ("游ゴシック", 12), ("游ゴシック", 16),
        ("Yu Gothic UI", 11), ("Yu Gothic UI", 12),
        ("メイリオ", 10), ("メイリオ", 11),
        ("Meiryo UI", 10), ("Meiryo UI", 11),
        ("Century", 9), ("Century", 11), ("Century", 12),
        ("Times New Roman", 10), ("Arial", 11), ("Arial", 12),
        ("Calibri", 11), ("Calibri", 12), ("Calibri", 14),
        ("Terminal", 14),
    ]
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    rows = []
    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        cell = ws.Range("A1")
        cell.WrapText = True
        print("%-16s %-4s %5s %4s %4s %4s %4s %4s %6s %6s" % (
            "face", "size", "excel", "asc", "des", "int", "ext", "K",
            "text", "allow"))
        for face, size in faces:
            sample = "あ" * 10
            tm, ink, resolved = font_facts(face, size, sample)
            if ink < 5 * int(round(size * 96.0 / 72.0)):
                sample = "n" * 10
                tm, ink, resolved = font_facts(face, size, sample)
            # the height this font asks a row for, on its own
            ws.Cells.Clear()
            ws.Rows(3).Font.Name = face
            ws.Rows(3).Font.Size = size
            ws.Rows(3).AutoFit()
            height = round(ws.Rows(3).Height / 0.75)
            # and the width a wrapped line of it keeps
            cell = ws.Range("A1")
            cell.WrapText = True
            cell.Value = sample
            cell.Font.Name = face
            cell.Font.Size = size
            fits = narrowest_fit(ws)
            rows.append({
                "face": face, "size": size, "height": height,
                "ascent": tm.tmAscent, "descent": tm.tmDescent,
                "internal": tm.tmInternalLeading,
                "external": tm.tmExternalLeading,
                "K": height - (tm.tmAscent + tm.tmDescent),
                "ink": ink, "allowance": fits - ink,
                "resolved": resolved,
            })
            r = rows[-1]
            print("%-16s %-4s %5d %4d %4d %4d %4d %4d %6d %6d%s" % (
                face, size, height, tm.tmAscent, tm.tmDescent,
                tm.tmInternalLeading, tm.tmExternalLeading, r["K"],
                ink, r["allowance"],
                "" if resolved == face else "  (SUB %s)" % resolved))
        wb.Close(False)
    finally:
        excel.Quit()
    with open(r"pipeline_data\com_measurements\xlsx_two_constants.json", "w",
              encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)
    print("\nK against the allowance:")
    pairs = {}
    for r in rows:
        pairs.setdefault((r["K"], r["allowance"]), []).append(
            "%s %s" % (r["face"], r["size"]))
    for (k, a), names in sorted(pairs.items()):
        print("   K=%d allowance=%d : %s" % (k, a, ", ".join(names)))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
