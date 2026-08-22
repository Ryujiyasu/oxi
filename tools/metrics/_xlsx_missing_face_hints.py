# -*- coding: utf-8 -*-
"""Does GDI find Excel's substitute when it is told what the file says?

`_xlsx_missing_face_panose.py` shows Excel replaces `AR P丸ゴシック体E` with
游ゴシック when the run carries `panose`, `pitchFamily` and `charset`, and with
something else when it carries only the name. Two of those three — the charset
and the pitch-and-family — have a place in a `LOGFONT`, which the renderer
currently fills with DEFAULT_CHARSET and DEFAULT_PITCH|FF_DONTCARE. This asks
GDI for the missing face under each combination and reports which face it
actually hands back.

    python tools\\metrics\\_xlsx_missing_face_hints.py
"""
import sys

import win32con
import win32ui

sys.stdout.reconfigure(encoding="utf-8")

MISSING = "AR P丸ゴシック体E"
SAID = "国「あ、い」。国ぁ"
POINTS = 12
PIXELS = -round(POINTS * 96 / 72)

# `charset="-128"` is SHIFT_JIS; `pitchFamily="50"` is 0x32 — family 3
# (FF_MODERN) with a variable pitch.
CHARSETS = [("default", win32con.DEFAULT_CHARSET), ("shiftjis", 128)]
FAMILIES = [
    ("dontcare|default", win32con.FF_DONTCARE | win32con.DEFAULT_PITCH),
    ("modern|variable", 0x32),
    ("swiss|variable", win32con.FF_SWISS | win32con.VARIABLE_PITCH),
    ("roman|variable", win32con.FF_ROMAN | win32con.VARIABLE_PITCH),
]

dc = win32ui.CreateDCFromHandle(win32ui.GetForegroundWindow().GetDC().GetSafeHdc())
print(f"{'charset':<10}{'pitch and family':<20}{'face used':<22}{'width':>7}")
for charset_name, charset in CHARSETS:
    for family_name, family in FAMILIES:
        font = win32ui.CreateFont(
            {
                "name": MISSING,
                "height": PIXELS,
                "charset": charset,
                "pitch and family": family,
                "quality": win32con.CLEARTYPE_QUALITY,
            }
        )
        held = dc.SelectObject(font)
        width = dc.GetTextExtent(SAID)[0]
        used = dc.GetTextFace()
        dc.SelectObject(held)
        print(f"{charset_name:<10}{family_name:<20}{used:<22}{width:>7}")

print()
print("for comparison, asked for by name:")
for face in ("游ゴシック", "ＭＳ Ｐゴシック", "ＭＳ ゴシック", "メイリオ", "Yu Gothic UI"):
    font = win32ui.CreateFont({"name": face, "height": PIXELS,
                               "charset": win32con.DEFAULT_CHARSET,
                               "quality": win32con.CLEARTYPE_QUALITY})
    held = dc.SelectObject(font)
    print(f"{'':<10}{'':<20}{dc.GetTextFace():<22}{dc.GetTextExtent(SAID)[0]:>7}")
    dc.SelectObject(held)
