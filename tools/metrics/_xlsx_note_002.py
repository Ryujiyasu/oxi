# -*- coding: utf-8 -*-
r"""Why does `002`'s note start its text nine pixels lower than a fresh one?

`_xlsx_note_lines.py` settled, on fourteen notes Excel itself made, that a
note's vertical inset changes nothing: written with top insets of 0, 3.6, 7.2
and 10.8 points — the file holds all four — Excel draws the first line of every
one of them in the same place. Yet `002`'s note, the same face at the same
size, has its first line fourteen pixels below its paper where a fresh one has
five. Acting on the fourteen-arm law without knowing why costs that workbook
0.0066, so the difference has to be named first.

Synthetic arms have said all they can, so this takes the real note and removes
one thing at a time — the same way the grey connector was pinned.

    python tools\metrics\_xlsx_note_002.py
    python tools\metrics\_xlsx_note_002.py --reuse
"""

from __future__ import annotations

import argparse
import re
import sys
import time
import zipfile
from pathlib import Path

import numpy as np
import win32com.client
from PIL import Image, ImageGrab

Image.MAX_IMAGE_PIXELS = None
REPO = Path(__file__).resolve().parents[2]
SOURCE = REPO / "tools" / "golden-test" / "documents" / "xlsx" / "b6a3a84180c9_002.xlsx"
SCRATCH = Path(r"C:\tmp\xlsx_note_002")
# The note under the glass, named by the anchor it hangs from. Its box's top
# rule lands at y=144 of the picture and its text at y=159.
ANCHOR = "64, 12, 0, 144, 97, 2, 2, 34"
WINDOW = (940, 1272, 130, 300)          # left, right, top, bottom


def shaped(vml: str, alter) -> str:
    """Rewrite the one `<v:shape>` that hangs from ANCHOR."""
    want = re.sub(r"\s+", "", ANCHOR)
    out = []
    for piece in re.split(r"(<v:shape\b.*?</v:shape>)", vml, flags=re.S):
        if piece.startswith("<v:shape"):
            found = re.search(r"<x:Anchor>\s*([^<]*)</x:Anchor>", piece)
            if found and re.sub(r"\s+", "", found.group(1)) == want:
                piece = alter(piece)
        out.append(piece)
    return "".join(out)


def set_inset(said: str | None):
    def alter(piece: str) -> str:
        piece = re.sub(r'\s+inset="[^"]*"', "", piece)
        if said is not None:
            piece = piece.replace("<v:textbox", f'<v:textbox inset="{said}"', 1)
        return piece

    return alter


def drop_fit(piece: str) -> str:
    return piece.replace(";mso-fit-shape-to-text:t", "").replace(
        "mso-fit-shape-to-text:t;", ""
    )


def two_lines(piece: str) -> str:
    """Leave the note holding less than its box, so nothing overflows."""
    return piece


ARMS: list[tuple[str, object]] = [
    ("as it stands", lambda one: one),
    ("no inset at all", set_inset(None)),
    ("inset top 0", set_inset("2.5mm,0,2.5mm,0")),
    ("inset top 7.2pt", set_inset("2.5mm,7.2pt,2.5mm,7.2pt")),
    ("inset top 20pt", set_inset("2.5mm,20pt,2.5mm,20pt")),
    ("no fit-shape-to-text", drop_fit),
    # Excel's own writer emits a PARTIAL list when only one margin is set —
    # `inset=",0"` — and the fourteen COM-made notes with partial lists all
    # drew in the same place whatever the number said. Whether the comma is
    # what Excel ignores is the last thing between the two readings.
    ("inset \",0\" partial", set_inset(",0")),
    ("inset \",20pt\" partial", set_inset(",20pt")),
]


def build(made: Path, alter) -> None:
    if made.exists():
        made.unlink()
    with zipfile.ZipFile(SOURCE) as was, zipfile.ZipFile(made, "w", zipfile.ZIP_DEFLATED) as now:
        for item in was.infolist():
            held = was.read(item.filename)
            if item.filename == "xl/drawings/vmlDrawing1.vml":
                held = shaped(held.decode("utf-8", "replace"), alter).encode("utf-8")
            now.writestr(item, held)


def shoot(made: Path, into: Path) -> bool:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(made))
    try:
        sheet = book.Worksheets(1)
        used = sheet.UsedRange
        for _ in range(8):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(1.4)
            held = ImageGrab.grabclipboard()
            if held is not None:
                held.save(into)
                return True
        return False
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()


def read(picture: Path) -> str:
    """The note's top rule, and where its lines of ink begin below it."""
    grey = np.asarray(Image.open(picture).convert("L")).astype(int)
    left, right, top, bottom = WINDOW
    if right > grey.shape[1] or bottom > grey.shape[0]:
        return "the window is off the picture"
    band = (grey[top:bottom, left:right] < 128).sum(axis=1)
    edge = next((i for i, v in enumerate(band) if v > 100), None)
    if edge is None:
        return "no note in the window"
    lit = [i for i, v in enumerate(band) if v > 8 and i > edge + 1]
    if not lit:
        return f"rule at {edge + top}, no text"
    runs, start, last = [], lit[0], lit[0]
    for i in lit[1:]:
        if i > last + 1:
            runs.append(start)
            start = i
        last = i
    runs.append(start)
    pitches = [runs[i + 1] - runs[i] for i in range(len(runs) - 1)]
    return (
        f"rule {edge + top:>4}  first +{runs[0] - edge:<3} lines {len(runs)}"
        f"  pitch " + (",".join(str(one) for one in pitches) if pitches else "-")
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    SCRATCH.mkdir(parents=True, exist_ok=True)
    for at, (name, alter) in enumerate(ARMS):
        made = SCRATCH / f"arm{at}.xlsx"
        shot = SCRATCH / f"arm{at}.png"
        if not args.reuse:
            build(made, alter)
            if not shoot(made, shot):
                print(f"  {name:<22} Excel would not hand over a picture")
                continue
        if not shot.exists():
            print(f"  {name:<22} no picture")
            continue
        print(f"  {name:<22} {read(shot)}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
