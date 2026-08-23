"""Does Oxi draw a shape's own outline where Excel draws it?

The corpus has sixteen `custGeom` shapes, but not one of them shows: the only
one on a rendered sheet sits left of the used range and is white on white. So
the outline code has nothing in the corpus to confirm it, and a feature no
picture exercises is a feature that is not known to work.

This authors the missing case: a freeform built through Excel's own
`BuildFreeform`, saved, drawn by Excel and by Oxi, and the two compared. A
closed filled triangle and an open stroked zig-zag, so both the fill path and
the stroke path are exercised.

Run: python tools/metrics/_xlsx_custgeom_repro.py
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
SCRATCH = Path(r"C:\tmp\xlsx_custgeom")


def author() -> Path:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "custgeom.xlsx"
    if made.exists():
        made.unlink()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        # A closed triangle, filled and ruled.
        build = sheet.Shapes.BuildFreeform(1, 60.0, 60.0)   # 1 = msoEditingAuto
        build.AddNodes(0, 1, 200.0, 60.0)                   # 0 = msoSegmentLine, 1 = msoEditingCorner
        build.AddNodes(0, 1, 130.0, 190.0)
        build.AddNodes(0, 1, 60.0, 60.0)
        made_shape = build.ConvertToShape()
        made_shape.Fill.ForeColor.RGB = 0x00CCFF   # BGR: a warm yellow
        made_shape.Line.ForeColor.RGB = 0x000000
        made_shape.Line.Weight = 1.5
        # An open zig-zag, stroked only.
        build = sheet.Shapes.BuildFreeform(1, 260.0, 60.0)
        for x, y in ((300.0, 190.0), (340.0, 60.0), (380.0, 190.0), (420.0, 60.0)):
            build.AddNodes(0, 1, x, y)
        open_shape = build.ConvertToShape()
        open_shape.Fill.Visible = False
        open_shape.Line.ForeColor.RGB = 0x000000
        open_shape.Line.Weight = 1.5
        sheet.Range("A1:Z40").Interior.Color = 0xFFFFFF
        book.SaveAs(str(made), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return made


def excel_picture(book_path: Path) -> Image.Image | None:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(book_path))
    try:
        sheet = book.Worksheets(1)
        for _ in range(6):
            try:
                sheet.Activate()
                sheet.Range("A1:Z40").CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.6)
                continue
            time.sleep(0.5)
            held = ImageGrab.grabclipboard()
            if held is not None:
                return held
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return None


def main() -> int:
    book_path = author()
    truth = excel_picture(book_path)
    if truth is None:
        print("Excel would not hand over a picture")
        return 1
    truth_png = SCRATCH / "excel.png"
    truth.save(truth_png)
    ours_png = SCRATCH / "oxi.png"
    subprocess.run(
        [str(RENDERER), str(book_path), str(ours_png), "96"],
        capture_output=True,
        env=dict(os.environ),
        timeout=900,
    )
    if not ours_png.exists():
        print("Oxi drew nothing")
        return 1
    a = np.asarray(Image.open(truth_png).convert("L"))
    b = np.asarray(Image.open(ours_png).convert("L"))
    high = min(a.shape[0], b.shape[0])
    wide = min(a.shape[1], b.shape[1])
    a, b = a[:high, :wide], b[:high, :wide]
    ink_a = a < 200
    ink_b = b < 200
    print(f"Excel ink {int(ink_a.sum())} px, Oxi ink {int(ink_b.sum())} px")
    print(f"in both {int((ink_a & ink_b).sum())}, Excel only {int((ink_a & ~ink_b).sum())},"
          f" Oxi only {int((ink_b & ~ink_a).sum())}")
    for name, mask in (("excel", ink_a), ("oxi", ink_b)):
        rows = np.where(mask.any(axis=1))[0]
        cols = np.where(mask.any(axis=0))[0]
        if len(rows):
            print(f"  {name} ink box rows {rows.min()}..{rows.max()} cols {cols.min()}..{cols.max()}")
        else:
            print(f"  {name} drew nothing")
    (SCRATCH / "_result.json").write_text(
        json.dumps({"excel_ink": int(ink_a.sum()), "oxi_ink": int(ink_b.sum())}, indent=1),
        encoding="utf-8",
    )
    print(f"pictures in {SCRATCH}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
