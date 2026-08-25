# -*- coding: utf-8 -*-
r"""What Excel puts in a CELL whose face it has not got, and how the charset
is spelt.

SX54 settled this for shapes: the name and the PANOSE make no difference, only
`charset="-128"` does, and it turns ＭＳ ゴシック into 游ゴシック. Two things
that reading never covered:

* it was measured in **shape** runs, and a cell's font is a different slot;
* the sweep spelt the Japanese charset one way. A file may write it **128** —
  the same byte, unsigned — and `sanko_tool` does exactly that.

The floor workbook says the two spellings are not the same: its red cell asks
for `AR P丸ゴシック体E` with `charset="128"` and Excel draws it 22 pixels wider
than GDI's own answer (ＭＳ Ｐゴシック), which is ＭＳ ゴシック's width, with the
full-width ： and １ of a fixed-pitch face. So this asks the cell question
directly, with every spelling in the sweep.

One workbook an arm — a name resolves once per document (SX54's second error),
so two arms in one book would wear each other's answer. Every installed
Japanese face rules the same workbook: those are installed, so nothing they do
can move the missing name's answer, and the arm is identified as the ruler
whose ink is the same ink.

    python tools\metrics\_xlsx_cell_missing_face.py
    python tools\metrics\_xlsx_cell_missing_face.py --reuse
"""

from __future__ import annotations

import argparse
import importlib
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import numpy as np
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
which = importlib.import_module("_xlsx_missing_face_which")

REPO = Path(__file__).resolve().parents[2]
SHOOTER = Path(__file__).resolve().parent / "_xlsx_screen_shot.ps1"
SCRATCH = Path(r"C:\tmp\xlsx_cell_missing_face")
MISSING = "AR P丸ゴシック体E"
# `sanko_tool`'s own line: the kanji are full width in every face, and the ：
# and the １ are what tell a fixed-pitch face from a proportional one.
WORDS = "手順１：調査票選択→ｱｲ「あ、い」。ｧ国Wij"
POINTS = 11.0
ROW_PT = 24.0
# What the arm's `<font>` carries beside the name, one thing at a time — and
# then every (name, family, charset) the corpus actually asks for and has not
# got, so the rule is read on the files that need it rather than on one name.
DRESSINGS = [
    (MISSING, "bare", ""),
    (MISSING, "charset 128", '<charset val="128"/>'),
    (MISSING, "charset -128", '<charset val="-128"/>'),
    (MISSING, "charset 0", '<charset val="0"/>'),
    (MISSING, "family 3", '<family val="3"/>'),
    (MISSING, "family 2", '<family val="2"/>'),
    (MISSING, "fam 3 + cs 128", '<family val="3"/><charset val="128"/>'),
    (MISSING, "fam 3 + cs -128", '<family val="3"/><charset val="-128"/>'),
    (MISSING, "fam 2 + cs 128", '<family val="2"/><charset val="128"/>'),
    (MISSING, "fam 1 + cs 128", '<family val="1"/><charset val="128"/>'),
    ("ＤＦ特太ゴシック体", "fam 3 + cs 128", '<family val="3"/><charset val="128"/>'),
    ("Nonesuch Gothic ZZ", "fam 3 + cs 128", '<family val="3"/><charset val="128"/>'),
    # The corpus's own asks. The spellings with spaces are near-misses of
    # installed faces, and GDI answers ＭＳ Ｐゴシック to a name nothing can
    # have — so a face that resolves THERE cannot be told from a missing one
    # by the device alone. Excel is the one being asked.
    ("MS P ゴシック", "fam 3 + cs 128", '<family val="3"/><charset val="128"/>'),
    ("MS　Pゴシック", "fam 3 + cs 128", '<family val="3"/><charset val="128"/>'),
    ("MS ゴシック", "fam 3 + cs 128", '<family val="3"/><charset val="128"/>'),
    ("MS PGothic", "bare", ""),
    ("明朝", "fam 1 + cs 128", '<family val="1"/><charset val="128"/>'),
    ("明朝", "fam 3 + cs 128", '<family val="3"/><charset val="128"/>'),
    ("HGSｺﾞｼｯｸ", "fam 3 + cs 128", '<family val="3"/><charset val="128"/>'),
    ("Arial Unicode MS", "fam 3 + cs 128", '<family val="3"/><charset val="128"/>'),
]
# The rulers are the Japanese faces; a Latin name may be answered with a Latin
# face, so the usual suspects stand beside them.
LATIN_RULERS = ["Arial", "Calibri", "Times New Roman", "Segoe UI", "Tahoma",
                "Courier New", "Verdana"]


def build(made: Path, asked: str, dressing: str, rulers: list[str]) -> None:
    """A workbook whose B2 asks for the missing face, with the rulers below."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    SCRATCH.mkdir(parents=True, exist_ok=True)
    plain = made.with_name("_plain_" + made.name)
    book = Workbook()
    sheet = book.active
    sheet.column_dimensions["B"].width = 40.0
    for at, face in enumerate([asked] + rulers, start=2):
        cell = sheet.cell(row=at, column=2, value=WORDS)
        cell.font = Font(name=face, size=POINTS)
        sheet.row_dimensions[at].height = ROW_PT
    book.save(plain)

    # openpyxl has no slot for a charset, so the arm's own `<font>` is dressed
    # in the part itself. Only the one that carries the missing name is
    # touched — the rulers keep whatever openpyxl wrote.
    made.unlink(missing_ok=True)
    with zipfile.ZipFile(plain) as source, \
            zipfile.ZipFile(made, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.namelist():
            body = source.read(item)
            if item == "xl/styles.xml" and dressing:
                text = body.decode("utf-8")
                text = re.sub(
                    r'(<font>(?:(?!</font>).)*?<name val="' + re.escape(asked) + r'"/>)',
                    r"\1" + dressing, text, count=1)
                body = text.encode("utf-8")
            out.writestr(item, body)
    plain.unlink(missing_ok=True)


def shoot(made: Path) -> Path:
    picture = made.with_suffix(".excel.png")
    picture.unlink(missing_ok=True)
    listing = SCRATCH / "_batch.txt"
    listing.write_text(f"{made.resolve()}\t{picture.resolve()}", encoding="utf-8")
    subprocess.run(["powershell", "-NoProfile", "-File", str(SHOOTER),
                    "-ListFile", str(listing.resolve())],
                   capture_output=True, text=True, encoding="utf-8",
                   errors="replace", timeout=600)
    listing.unlink(missing_ok=True)
    return picture


def bands(picture: np.ndarray, count: int) -> list[np.ndarray | None]:
    """Each cell's own ink, cropped to itself.

    The rows are all one height, so the band is arithmetic; the reading skips
    two pixels at each edge, where the sheet's own gridline is ink.
    """
    tall = round(ROW_PT * 96 / 72)
    held = []
    for at in range(count):
        top = at * tall
        band = picture[top + 2:top + tall - 2, 2:]
        held.append(which.ink_of(band, 0, band.shape[0]))
    return held


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    rulers = which.japanese_faces() + LATIN_RULERS
    print(f"  {WORDS!r} at {POINTS}pt, against {len(rulers)} installed faces")
    print("  asked                     dressing          Excel draws")
    for at, (asked, name, dressing) in enumerate(DRESSINGS):
        made = SCRATCH / f"cell_{at:02d}.xlsx"
        build(made, asked, dressing, rulers)
        picture = made.with_suffix(".excel.png") if args.reuse else shoot(made)
        if not picture.exists():
            print(f"  {asked:<24}  {name:<16}  Excel gave no picture")
            continue
        truth = np.asarray(Image.open(picture).convert("L"))
        inks = bands(truth, len(rulers) + 1)
        drawn = inks[0]
        if inks[0] is None:
            print(f"  {asked:<24}  {name:<16}  nothing to read")
            continue
        scored = []
        for face, ink in zip(rulers, inks[1:]):
            told = which.unlike(drawn, ink)
            if told is not None:
                scored.append((told[0], face, told[1]))
        scored.sort()
        if not scored:
            print(f"  {asked:<24}  {name:<16}  no ruler to read")
            continue
        best = scored[0]
        near = ", ".join(f"{face} {off}" for off, face, _ in scored[1:3])
        told = f"{best[1]}" if best[0] == 0 else f"none ({best[1]} off by {best[0]})"
        print(f"  {asked:<24}  {name:<16}  {told:<22} /{best[2]:<5} (next: {near})")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
