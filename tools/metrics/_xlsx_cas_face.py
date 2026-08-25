# -*- coding: utf-8 -*-
r"""Which face is Excel drawing the `cas-` title with?

The ten `cas-r*` workbooks ask for `ＤＦ特太ゴシック体`, which this machine has
not got, and 78% of each book's differing pixels are in that one title. We
answer the missing name with 游ゴシック (SX54) and our line comes out two pixels
wider, so Excel is answering with something else.

Reading it off the picture does not work. Comparing Excel's ink to a GDI render
of a candidate compares two RASTERISERS: the ClearType phase differs, the `（`
splits into two blobs in one and not the other, and every candidate scores
"near" on ink widths within a pixel. Three attempts, no answer.

So the rulers are drawn by EXCEL ITSELF, inside a copy of the workbook that is
asking the question — which is also the only way to keep the per-document
resolution SX54 found. Every installed Japanese face gets a shape holding the
same title at the same size, and the answer is the ruler whose ink is the same
ink.

    python tools\metrics\_xlsx_cas_face.py
    python tools\metrics\_xlsx_cas_face.py --reuse
"""

from __future__ import annotations

import argparse
import importlib
import shutil
import sys
import time
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
SCRATCH = Path(r"C:\tmp\xlsx_cas_face")
BOOK = REPO / "tools" / "golden-test" / "documents" / "xlsx" / "3e2edebd2a0c_cas-r02gassan-4hyou.xlsx"
TITLE = "内閣所管（合算）"
POINTS = 14.0
HIGH, WIDE = 26.0, 260.0
GAP = 4.0
sys.path.insert(0, str(Path(__file__).resolve().parent))
which = importlib.import_module("_xlsx_missing_face_which")


def build(made: Path, faces: list[tuple[str, str, bool, bool, float]]) -> list[float] | None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    shutil.copy(BOOK, made)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(made))
    tops = []
    try:
        sheet = book.Worksheets(1)
        # Well below the sheet's own content, so nothing of the workbook's is
        # in the band — and on a white ground, like the title's own.
        at = 44.0
        for _label, face, centred, wraps, wide in faces:
            shape = sheet.Shapes.AddShape(1, 18.0, at, wide, HIGH)
            frame = shape.TextFrame2
            frame.WordWrap = wraps
            frame.AutoSize = 0
            frame.VerticalAnchor = 1
            frame.TextRange.Text = TITLE
            frame.TextRange.Font.Size = POINTS
            frame.TextRange.Font.Name = face
            try:
                frame.TextRange.Font.NameFarEast = face
            except Exception:
                pass
            frame.TextRange.Font.Fill.ForeColor.RGB = 0
            frame.TextRange.ParagraphFormat.Alignment = (
                3 if centred == "right" else 2 if centred else 1)
            shape.Fill.Visible = True
            shape.Fill.ForeColor.RGB = 0xFFFFFF
            shape.Line.Visible = False
            try:
                shape.Shadow.Visible = False
            except Exception:
                pass
            tops.append(at)
            at += HIGH + GAP
        book.Save()
    finally:
        book.Close(SaveChanges=True)
        excel.Quit()
    return tops


def dress(made: Path) -> None:
    """Put the title's own dressing on the rulers COM built bare.

    The file's own title states `pitchFamily="49" charset="-128"` on its
    typefaces and a COM-built run states neither, which is the last thing that
    differs between them. Only the bare ones are touched, so the title keeps
    what it had.
    """
    import re
    import shutil
    import zipfile

    beside = made.with_name("_dressed_" + made.name)
    with zipfile.ZipFile(made) as source,             zipfile.ZipFile(beside, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.namelist():
            body = source.read(item)
            if item.startswith("xl/drawings/drawing") and item.endswith(".xml"):
                text = body.decode("utf-8")
                # COM writes a `panose` of its own for a face it knows, so
                # matching only `<a:latin typeface="…"/>` patches nothing.
                # Anything that has no pitchFamily yet gets the title's.
                def wear(found: "re.Match[str]") -> str:
                    held = found.group(0)
                    if "pitchFamily" in held:
                        return held
                    return held[:-2] + ' pitchFamily="49" charset="-128"/>'

                text = re.sub(r'<a:(?:latin|ea)[^>]*/>', wear, text)
                body = text.encode("utf-8")
            out.writestr(item, body)
    shutil.move(str(beside), str(made))


def shoot(made: Path) -> bool:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(made))
    try:
        sheet = book.Worksheets(1)
        for _ in range(10):
            try:
                sheet.Activate()
                sheet.Range("A1:H24").CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(1.2)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(SCRATCH / "excel.png")
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def ink_at(picture: np.ndarray, top: float, wide: float = WIDE) -> np.ndarray | None:
    # The ruler's own box only: the same row band carries the sheet's content
    # either side of it, and taking the whole width reads that as the ruler's
    # ink (691 pixels for a line that is 136).
    left = round(18.0 * 96 / 72) + 2
    right = round((18.0 + wide) * 96 / 72) - 2
    band = picture[round(top * 96 / 72) + 2:round((top + HIGH) * 96 / 72) - 2,
                   left:right] < 128
    rows = np.flatnonzero(band.any(axis=1))
    columns = np.flatnonzero(band.any(axis=0))
    if not rows.size or not columns.size:
        return None
    return band[rows[0]:rows[-1] + 1, columns[0]:columns[-1] + 1]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    parser.add_argument("--dress", action="store_true",
                        help="put pitchFamily/charset on the rulers COM built bare")
    args = parser.parse_args()
    # The whole installed list would need 2000 points of sheet to stand in,
    # and the capture has to stay inside the book's own used range (A1:H24) or
    # Excel hands back a picture of a different shape. These are the faces the
    # ink-width sweep left standing.
    # (label, face, centred, wraps, box width in points). The first arms are
    # the faces; the rest put the TITLE's own properties back one at a time,
    # because the missing name and 游ゴシック draw the same ink (136) and the
    # title's is 134 — so what differs is the shape, not the face.
    DF = "ＤＦ特太ゴシック体"
    faces = [("df", DF, False, False, 260.0),
             ("yu", "游ゴシック", False, False, 260.0),
             ("df centred", DF, True, False, 260.0),
             ("df wraps", DF, False, True, 260.0),
             ("df narrow", DF, False, False, 167.0),
             ("df all three", DF, True, True, 167.0),
             ("meiryo", "メイリオ", False, False, 260.0),
             ("ms p", "ＭＳ Ｐゴシック", False, False, 260.0),
             # A left-aligned and a right-aligned twin of the same dressed run:
             # the gap between where their lines START is `room - width`, so
             # Excel's own width for the line falls out without trusting any of
             # our own arithmetic.
             ("left twin", DF, False, False, 260.0),
             ("right twin", DF, "right", False, 260.0),
             # The same pair for a LOOSE run (an installed name, no give-backs
             # in eight glyphs): if Excel's width is the rounded EXACT sum it
             # reads 149 here too, where the drawn steps sum to 152.
             ("yu left twin", "游ゴシック", False, False, 260.0),
             ("yu right twin", "游ゴシック", "right", False, 260.0),
]
    # The missing name ITSELF is the control: if its ruler draws the title's
    # own ink then the substitution is behaving as it does in the title, and
    # what differs from plain 游ゴシック is the substitution — not the shape.
    made = SCRATCH / "casface.xlsx"
    if args.reuse:
        tops = [44.0 + at * (HIGH + GAP) for at in range(len(faces))]
    else:
        tops = build(made, faces)
        if tops is None:
            print("  Excel would not build the rulers")
            return 1
        if args.dress:
            dress(made)
        if not shoot(made):
            print("  Excel would not hand over a picture")
            return 1
    shot = np.asarray(Image.open(SCRATCH / "excel.png").convert("L"))
    # The title itself, at the top of the sheet — and only its own box: the
    # same row band carries the sheet's own headings further right, and taking
    # the whole width reads them as part of the title (691 pixels of "ink" for
    # a line that is 134).
    band = shot[12:36, 45:200] < 128
    # Only the title's own corner of the sheet: the same rows carry the
    # workbook's headings further right, and the whole width reads as 691
    # pixels of "ink" for a line that is 136. The box around the title is
    # drawn in a pale theme colour and does not reach the threshold.
    rows = np.flatnonzero(band.any(axis=1))
    columns = np.flatnonzero(band.any(axis=0))
    title = band[rows[0]:rows[-1] + 1, columns[0]:columns[-1] + 1]
    print(f"  the title's own ink {title.shape}")
    scored = []
    for (label, _face, _centred, _wraps, wide), top in zip(faces, tops):
        got = ink_at(shot, top, wide)
        told = which.unlike(title, got)
        if told:
            scored.append((told[0], label, got.shape))
    scored.sort()
    twins, loose = {}, {}
    for (label, _face, _centred, _wraps, wide), top in zip(faces, tops):
        if "twin" not in label:
            continue
        got = ink_at(shot, top, wide)
        if got is None:
            continue
        # Where the ink starts inside the ruler's own box.
        left = round(18.0 * 96 / 72) + 2
        band = shot[round(top * 96 / 72) + 2:round((top + HIGH) * 96 / 72) - 2,
                    left:round((18.0 + wide) * 96 / 72) - 2] < 128
        columns = np.flatnonzero(band.any(axis=0))
        (loose if label.startswith("yu") else twins)[label] = int(columns[0])
    if len(twins) == 2:
        print(f"  dressed twins {twins}; room - width = "
              f"{twins['right twin'] - twins['left twin']}")
    if len(loose) == 2:
        print(f"  loose twins   {loose}; room - width = "
              f"{loose['yu right twin'] - loose['yu left twin']}")
    for off, face, shape in scored[:8]:
        print(f"    {face:<24} unlike {off:>5}  ink {shape}"
              f"{'   <== the same ink' if off == 0 else ''}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
