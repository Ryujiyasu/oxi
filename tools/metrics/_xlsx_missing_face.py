# -*- coding: utf-8 -*-
"""What a face this machine does not have comes out as.

`sanko_tool` asks its callouts for `AR P丸ゴシック体E`, which is not installed.
Excel draws the text at a full 16px a character for a 12pt run — a full-width
fallback — while GDI, asked for the same name, hands back something that
squeezes the punctuation. This measures both so the substitution can be
matched rather than guessed.

    python tools\\metrics\\_xlsx_missing_face.py
"""
import sys

import win32con
import win32ui

sys.stdout.reconfigure(encoding="utf-8")

SAID = "確認したい品目アイテムについて、「調査票番号」"
FACES = [
    "AR P丸ゴシック体E",  # what the shape asks for, and this machine has not
    "ＭＳ ゴシック",
    "ＭＳ Ｐゴシック",
    "ＭＳ 明朝",
    "ＭＳ Ｐ明朝",
    "メイリオ",
    "Meiryo UI",
    "游ゴシック",
    "Yu Gothic UI",
    "MS UI Gothic",
]
POINTS = 12
PIXELS = -round(POINTS * 96 / 72)

dc = win32ui.CreateDCFromHandle(win32ui.GetForegroundWindow().GetDC().GetSafeHdc())
print(f"{'face asked for':<22}{'face used':<22}{'width':>7}{'per char':>10}")
for face in FACES:
    font = win32ui.CreateFont(
        {
            "name": face,
            "height": PIXELS,
            "charset": win32con.DEFAULT_CHARSET,
            "quality": win32con.CLEARTYPE_QUALITY,
        }
    )
    held = dc.SelectObject(font)
    width = dc.GetTextExtent(SAID)[0]
    used = dc.GetTextFace()
    dc.SelectObject(held)
    print(f"{face:<22}{used:<22}{width:>7}{width / len(SAID):>10.2f}")
