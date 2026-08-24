"""Does a run's <rPr> replace the cell's font, or only override what it names?

`glossary_05` puts a bold 12pt Yu Gothic UI font on `C8` and fills it from a
shared string of two runs: the first carries no `<rPr>` at all, the second
carries one that names a size and a face but NOT `<b/>`. Excel draws the first
bold and the second regular; Oxi draws both bold, because it computes
`run.bold || cell.bold`. Nine of that book's seventeen strings carry runs.

If `<rPr>` REPLACES the cell font, a run whose `<rPr>` says only `<b/>` must
come out in the default face and size — not the cell's 20pt. If it only
OVERRIDES, that run stays 20pt. The two are told apart by the ink's height,
which is why the cell font is made large here.

The runs are written into `sharedStrings.xml` by hand: COM always emits a full
`<rPr>` and so cannot pose the question.

Run: python tools/metrics/_xlsx_run_props.py
"""

from __future__ import annotations

import re
import sys
import time
import zipfile
from pathlib import Path

import numpy as np
import win32com.client
from PIL import ImageGrab

SCRATCH = Path(r"C:\tmp\xlsx_run_props")
CELL_FACE, CELL_SIZE = "ＭＳ ゴシック", 20.0
# Each arm: what the second run's <rPr> holds, and what it is asking.
ARMS = [
    ("", "no rPr at all"),
    ("<sz val=\"20\"/><rFont val=\"ＭＳ ゴシック\"/>", "size and face, no b"),
    ("<b/>", "b alone"),
    ("<sz val=\"10\"/>", "size alone"),
    ("<b/><sz val=\"20\"/><rFont val=\"ＭＳ ゴシック\"/>", "b, size and face"),
]
HEAD, TAIL = "アアアア", "イイイイ"


def seed() -> Path:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "seed.xlsx"
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:D20").Interior.Color = 0xFFFFFF
        sheet.Columns("B").ColumnWidth = 30.0
        for at in range(2, 2 + len(ARMS)):
            cell = sheet.Cells(at, 2)
            cell.Value = HEAD + TAIL
            cell.Font.Name = CELL_FACE
            cell.Font.Size = CELL_SIZE
            cell.Font.Bold = True
            sheet.Rows(at).RowHeight = 34.0
        book.SaveAs(str(made), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return made


def written(source: Path) -> Path:
    out = SCRATCH / "runs.xlsx"
    if out.exists():
        out.unlink()
    held = zipfile.ZipFile(source)
    pieces = []
    for props, _ in ARMS:
        opened = f"<rPr>{props}</rPr>" if props else ""
        pieces.append(
            f"<si><r><t xml:space=\"preserve\">{HEAD}</t></r>"
            f"<r>{opened}<t xml:space=\"preserve\">{TAIL}</t></r></si>"
        )
    fresh = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             f'count="{len(ARMS)}" uniqueCount="{len(ARMS)}">' + "".join(pieces) + "</sst>")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as writing:
        for item in held.infolist():
            raw = held.read(item.filename)
            if item.filename == "xl/sharedStrings.xml":
                raw = fresh.encode("utf-8")
            elif item.filename.startswith("xl/worksheets/sheet"):
                text = raw.decode("utf-8")
                # Point each row's cell at its own string.
                for at in range(len(ARMS)):
                    text = re.sub(rf'(<c r="B{at + 2}"[^>]*t="s"[^>]*><v>)\d+(</v>)',
                                  rf'\g<1>{at}\g<2>', text)
                raw = text.encode("utf-8")
            writing.writestr(item, raw)
    return out


def main() -> int:
    book_path = written(seed())
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(book_path))
    try:
        sheet = book.Worksheets(1)
        used = sheet.Range(sheet.Cells(2, 2), sheet.Cells(1 + len(ARMS), 2))
        for _ in range(10):
            try:
                sheet.Activate()
                used.CopyPicture(Appearance=1, Format=2)
            except Exception:
                time.sleep(0.8)
                continue
            time.sleep(0.6)
            shot = ImageGrab.grabclipboard()
            if shot is not None:
                break
        else:
            print("Excel would not hand over a picture")
            return 1
        shot.save(SCRATCH / "shot.png")
        grey = np.asarray(shot.convert("L"))
        base = round(sheet.Cells(2, 2).Top * 96 / 72)
        print(f"  cell font: {CELL_FACE} {CELL_SIZE:.0f}pt BOLD;"
              f" first run bare, second run as named")
        print("  what the second run's rPr holds     run 1 ink/high   run 2 ink/high")
        for at in range(len(ARMS)):
            top = round(sheet.Cells(at + 2, 2).Top * 96 / 72) - base
            high = round(sheet.Cells(at + 2, 2).Height * 96 / 72)
            left = 0
            wide = round(sheet.Cells(at + 2, 2).Width * 96 / 72)
            block = grey[top:top + high, left:left + wide] < 140
            lit = np.where(block.any(axis=0))[0]
            if not len(lit):
                print(f"  {ARMS[at][1]:<34}  no ink")
                continue
            # Split at the widest gap: the two runs are four letters each.
            gaps = np.where(np.diff(lit) > 3)[0]
            cut = lit[gaps[len(gaps) // 2]] if len(gaps) else (lit.min() + lit.max()) // 2
            out = []
            for a, b in ((lit.min(), cut + 1), (cut + 1, lit.max() + 1)):
                part = block[:, a:b]
                rows = np.where(part.any(axis=1))[0]
                out.append((int(part.sum()), int(rows.max() - rows.min() + 1)
                            if len(rows) else 0))
            print(f"  {ARMS[at][1]:<34}  {out[0][0]:>5} / {out[0][1]:<3}"
                  f"      {out[1][0]:>5} / {out[1][1]:<3}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
