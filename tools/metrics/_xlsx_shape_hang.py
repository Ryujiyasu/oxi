"""Does a shape let a closing 約物 hang past the end of its line?

`dc4fcff7f5f8_001`'s panel ends a line on 「ご確認ください。」. Oxi fits every
character up to 「い」 and then, because 「。」 may not start a line, kinsoku drags
「い」 down with it — two characters onto a line of their own. Excel keeps all of
them. Hanging punctuation (ぶら下げ) would do exactly that.

The arm: a box whose room is swept across the point where the line's last
character stops fitting, over eight last characters. If a line holds together
after the plain-kana line has broken, that character hangs.

ANSWER (ＭＳ ゴシック 12pt, 18-character body, one em = 16px): 「。」 and 「、」
hold all the way down to room = the body alone (short 16) and break only at 17
— they hang their whole em. 「」」「）」「！」「ゃ」「あ」 all break the moment the
room is 2px short: they do not hang. 「)」 looks like it hangs but does not — it
is half-width, so body + ) = 296px and it simply fits until the room drops
below that. Reading a half-width character on an em-sized sweep is the trap
here: state the last character's OWN width before calling anything a hang.

So the hanging set is the two 句読点 and nothing else.

Run: python tools/metrics/_xlsx_shape_hang.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_shape_hang")
FACE = "ＭＳ ゴシック"   # a face whose every glyph is one em: the arithmetic is exact
SIZE = 12.0
EM = SIZE * 96 / 72     # 16 px a character
TOP_PT = 100.0
LEFT_PT = 60.0
HIGH_PT = 60.0
BODY = "いろはにほへとちりぬるをわかよたれそ"   # 18 characters


def bands(image, left: int, right: int) -> list[tuple[int, int]]:
    grey = np.asarray(image.convert("L"))[:, left:right]
    lit = (grey < 120).sum(axis=1)
    out: list[tuple[int, int]] = []
    start = None
    for row, count in enumerate(lit):
        if count > 1 and start is None:
            start = row
        if count <= 1 and start is not None:
            out.append((start, row - 1))
            start = None
    if start is not None:
        out.append((start, len(lit) - 1))
    return out


def picture(sheet):
    for _ in range(6):
        try:
            sheet.Activate()
            sheet.Range("A1:Z60").CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.4)
        held = ImageGrab.grabclipboard()
        if held is not None:
            return held
    return None


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z60").Interior.Color = 0xFFFFFF
        print(f"{FACE} {SIZE:.0f}pt, one em = {EM:.2f}px, {len(BODY)} characters + a last one")
        print("room is the box less its two 7.2pt margins; 'n' is how many ink lines came out")
        # Which characters hang, and which only refuse to start a line.
        for last, name in (("。", "kuten"), ("、", "touten"), ("」", "close-kagi"),
                           ("）", "close-paren-wide"), (")", "close-paren"),
                           ("！", "bang"), ("ゃ", "small-ya"), ("あ", "plain")):
            print(f"  last character 「{last}」")
            for short in (0, 2, 8, 14, 16, 18):
                # Room enough for the body plus the last character, less `short`
                # pixels: the crossing sits at short = 0.
                room_px = (len(BODY) + 1) * EM - short
                wide_pt = (room_px + 2 * 9.6) * 72 / 96
                shape = sheet.Shapes.AddShape(1, LEFT_PT, TOP_PT, wide_pt, HIGH_PT)
                try:
                    frame = shape.TextFrame2
                    frame.WordWrap = True
                    frame.AutoSize = 0
                    frame.VerticalAnchor = 1
                    frame.TextRange.Text = BODY + last
                    frame.TextRange.Font.Size = SIZE
                    frame.TextRange.Font.Name = FACE
                    frame.TextRange.Font.Fill.ForeColor.RGB = 0
                    shape.Fill.Visible = False
                    shape.Line.Visible = False
                    held = picture(sheet)
                    if held is None:
                        print("    Excel would not hand over a picture")
                        continue
                    window = (round(LEFT_PT * 96 / 72) + 4,
                              round((LEFT_PT + wide_pt) * 96 / 72) + 40)
                    found = bands(held, *window)
                    print(f"    room {room_px:6.1f} (short {short:2d})  n {len(found)}")
                finally:
                    shape.Delete()
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
