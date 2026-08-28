# -*- coding: utf-8 -*-
r"""What each face says about itself, per em, at the size GDI is asked for.

The shape pitch is `tmHeight` scaled to the size and multiplied by 1.3 for a
Japanese face, and that much is settled to a hundredth of a pixel. What is not
settled is the FRACTION the block of lines starts on: the renderer starts it
on a whole pixel and Excel does not, and to model the fraction one has to know
which of the face's own numbers carries it.

Measured at 2048 pixels so the ratios are exact, and again at the asked-for
pixel size so the whole-pixel numbers GDI would actually use are visible.

    python tools\metrics\_xlsx_face_metrics.py
"""

from __future__ import annotations

import sys

import win32con
import win32gui
import win32ui

FACES = ["游ゴシック", "Yu Gothic UI", "メイリオ", "ＭＳ Ｐゴシック"]
SIZES = [9.0, 10.0, 11.0, 12.0, 14.0, 16.0, 18.0, 20.0]
BIG = 2048


def metrics(face: str, pixels: int) -> dict:
    screen = win32gui.GetDC(0)
    dc = win32ui.CreateDCFromHandle(screen)
    font = win32ui.CreateFont({
        "name": face,
        "height": -pixels,
        "weight": win32con.FW_NORMAL,
        "charset": win32con.DEFAULT_CHARSET,
    })
    old = dc.SelectObject(font)
    told = dc.GetTextMetrics()
    dc.SelectObject(old)
    win32gui.ReleaseDC(0, screen)
    return told


def main() -> int:
    print(f"{'face':<14}{'height/em':>11}{'ascent/em':>11}{'descent/em':>11}"
          f"{'intlead/em':>11}")
    for face in FACES:
        told = metrics(face, BIG)
        print(f"{face:<14}{told['tmHeight'] / BIG:>11.5f}{told['tmAscent'] / BIG:>11.5f}"
              f"{told['tmDescent'] / BIG:>11.5f}{told['tmInternalLeading'] / BIG:>11.5f}")

    print(f"\n{'face':<14}{'pt':>6}{'em px':>8}{'height':>8}{'ascent':>8}"
          f"{'descent':>8}{'exact height':>14}{'exact ascent':>14}")
    for face in FACES:
        ratio = metrics(face, BIG)
        for points in SIZES:
            em = points * 96.0 / 72.0
            told = metrics(face, round(em))
            print(f"{face:<14}{points:>6}{em:>8.3f}{told['tmHeight']:>8}"
                  f"{told['tmAscent']:>8}{told['tmDescent']:>8}"
                  f"{ratio['tmHeight'] / BIG * em:>14.4f}"
                  f"{ratio['tmAscent'] / BIG * em:>14.4f}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
