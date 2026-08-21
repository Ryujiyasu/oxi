# -*- coding: utf-8 -*-
"""Ask PowerPoint which font it resolves for each run of a slide.

The inheritance chain in the XML can be read three ways when a shape is both a
placeholder and a text box; PowerPoint's own answer settles it. Reports
Font.Name / Size / Bold per run, plus the shape's placeholder identity.

Usage: python tools/metrics/read_pptx_runfont.py <deck.pptx> <slide-number>
"""
from __future__ import annotations

import os
import sys

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    path = os.path.abspath(sys.argv[1])
    index = int(sys.argv[2])
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(path, WithWindow=False)
    try:
        slide = pres.Slides(index)
        for shape in slide.Shapes:
            try:
                if not shape.HasTextFrame or not shape.TextFrame.HasText:
                    continue
            except Exception:
                continue
            ph = ""
            try:
                ph = f" ph(type={shape.PlaceholderFormat.Type})"
            except Exception:
                pass
            print(f"\n[{shape.Name}]{ph}")
            rng = shape.TextFrame.TextRange
            for i in range(1, rng.Runs().Count + 1):
                run = rng.Runs(i)
                text = str(run.Text)[:44].replace("\r", " ")
                f = run.Font
                print(f"   run{i}: {f.Name!r} sz={f.Size} bold={f.Bold} "
                      f"text={text!r}")
    finally:
        pres.Close()


if __name__ == "__main__":
    main()
