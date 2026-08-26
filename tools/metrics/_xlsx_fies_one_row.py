# -*- coding: utf-8 -*-
r"""One row of `fies_t2`, photographed as it is and then with one thing changed.

The scratch reproduction says a merged, distributed, 12pt ＭＳ 明朝 label in a
19px row sits at offset 3. The book's own G131 sits at 2, and its row grid
agrees with ours everywhere (best dy is 0 over three separate windows), so the
pixel is real and some input is still unaccounted for.

Rather than keep adding to the scratch, this takes the book apart: it
photographs G131:L131 as it stands, then again after removing the merge, then
after moving the alignment off distributed. Whichever removal takes the ink
from 2 to 3 names the input.

    python tools\metrics\_xlsx_fies_one_row.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
BOOK = REPO / "tools" / "golden-test" / "documents" / "xlsx" / "d1eb05860dd5_fies_t2.xlsx"
SCRATCH = Path(r"C:\tmp\xlsx_fies_one_row")
WHERE = "G131:L131"


def shot(sheet, name: str) -> tuple[int, int] | None:
    for _ in range(10):
        try:
            sheet.Activate()
            sheet.Range(WHERE).CopyPicture(Appearance=1, Format=2)
        except Exception:
            time.sleep(0.6)
            continue
        time.sleep(0.8)
        held = ImageGrab.grabclipboard()
        if held is None:
            continue
        held.save(SCRATCH / f"{name}.png")
        dark = np.asarray(held.convert("L")) < 140
        rows = [r for r in range(dark.shape[0]) if dark[r].sum() >= 4]
        return (rows[0], rows[-1], dark.shape[0]) if rows else None
    return None


def main() -> int:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(BOOK), ReadOnly=True)
    try:
        sheet = book.Worksheets(1)
        told = {}
        told["as it stands (distributed)"] = shot(sheet, "stands")
        for name, how in (("left", -4131), ("centre", -4108), ("right", -4152),
                          ("justify", -4130), ("distributed again", -4117)):
            sheet.Range("G131").HorizontalAlignment = how
            told[name] = shot(sheet, name.replace(" ", "_"))
        sheet.Range("G131").HorizontalAlignment = -4117
        sheet.Range("G131:K131").UnMerge()
        told["unmerged, distributed"] = shot(sheet, "unmerged")

        print(f"  {BOOK.name}  {WHERE}")
        for name, held in told.items():
            if held is None:
                print(f"    {name:<18}  nothing to read")
                continue
            top, foot, tall = held
            print(f"    {name:<18}  ink {top}..{foot} in a {tall}px picture")
        return 0
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
