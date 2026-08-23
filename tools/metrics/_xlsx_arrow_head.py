"""How big is a line's arrowhead, and what shape?

`glossary_05` (the second-lowest scoring workbook) and three `zuhyo` books
draw flowcharts whose connectors carry arrowheads; Oxi draws the lines bare.
76 heads across seven workbooks — `triangle` 64 and `arrow` 12, every one of
them at the default size or `med`, which is the same thing.

The heads are read off Excel's own picture: a horizontal line with a head on
its tail, swept over line width, so how the head scales with the rule is
measured rather than assumed. What is reported is the head's length along the
line and its half-width across it.

Run: python tools/metrics/_xlsx_arrow_head.py
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

SCRATCH = Path(r"C:\tmp\xlsx_arrow")
LEFT_PT, TOP_PT = 60.0, 60.0
LONG_PT = 200.0
# Line width in points, the head's kind, and its stated size.
# The corpus's arrowed lines are almost all 0.75pt (9525 EMU = 1px), with a
# few at 1pt and 1.75pt, so the sweep is over the widths that actually occur
# rather than over round numbers.
ARMS = [
    (0.75, "triangle", None),
    (1.0, "triangle", None),
    (1.75, "triangle", None),
    (0.75, "arrow", None),
    (1.0, "arrow", None),
    (1.75, "arrow", None),
    (0.75, "triangle", ("med", "med")),
    (0.75, "stealth", None),
    (0.75, "oval", None),
]


def seed(weight: float) -> Path:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / f"seed_{weight:.2f}.xlsx"
    if made.exists():
        return made
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z40").Interior.Color = 0xFFFFFF
        line = sheet.Shapes.AddLine(LEFT_PT, TOP_PT, LEFT_PT + LONG_PT, TOP_PT)
        line.Line.ForeColor.RGB = 0x000000
        line.Line.Weight = weight
        book.SaveAs(str(made), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return made


def with_head(source: Path, kind: str, size) -> Path:
    out = SCRATCH / f"{source.stem}_{kind}_{size[0] if size else 'default'}.xlsx"
    if out.exists():
        out.unlink()
    held = zipfile.ZipFile(source)
    stated = f' w="{size[0]}" len="{size[1]}"' if size else ""
    tail = f'<a:tailEnd type="{kind}"{stated}/>'
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as writing:
        for item in held.infolist():
            raw = held.read(item.filename)
            if item.filename.startswith("xl/drawings/") and item.filename.endswith(".xml"):
                text = raw.decode("utf-8")
                # The rule element closes itself or holds children; put the
                # tail inside it either way.
                text = re.sub(r"(<a:ln\b[^>]*)/>", lambda m: m.group(1) + ">" + tail + "</a:ln>", text)
                if tail not in text:
                    text = re.sub(r"(</a:ln>)", lambda m: tail + m.group(1), text, count=1)
                raw = text.encode("utf-8")
            writing.writestr(item, raw)
    return out


def picture(sheet):
    for _ in range(8):
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
    return None


def ink_of(image):
    top = round(TOP_PT * 96 / 72)
    left = round(LEFT_PT * 96 / 72)
    right = left + round(LONG_PT * 96 / 72)
    grey = np.asarray(image.convert("L"))[top - 40:top + 40, left - 10:right + 40]
    return grey < 140


def measure(bare, headed):
    """The head, as the ink one picture has and the other has not.

    Every attempt to find the head inside a single picture came unstuck on the
    same thing: a head is thinnest at its tip and thickest at its base, so
    both "walk back from the end while it is fat" and "keep only what stands
    off the rule" measure some other quantity. Differencing against the same
    line drawn WITHOUT a head leaves the head and nothing else.
    """
    grew = headed & ~bare
    cols = np.where(grew.any(axis=0))[0]
    rows = np.where(grew.any(axis=1))[0]
    if not len(cols):
        return None
    end = np.where(bare.any(axis=0))[0].max()
    return {
        "length": int(cols.max() - cols.min() + 1),
        "across": int(rows.max() - rows.min() + 1),
        "past the end": int(cols.max() - end),
    }


def main() -> int:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    bare_ink: dict[float, object] = {}
    print("  weight  kind      size      head length  across  past the end   as multiples of the rule")
    try:
        for weight, kind, size in ARMS:
            source = seed(weight)
            if weight not in bare_ink:
                book = excel.Workbooks.Open(str(source))
                try:
                    shown = picture(book.Worksheets(1))
                    bare_ink[weight] = ink_of(shown) if shown is not None else None
                finally:
                    book.Close(SaveChanges=False)
            book_path = with_head(source, kind, size)
            book = excel.Workbooks.Open(str(book_path))
            try:
                held = picture(book.Worksheets(1))
                if held is None or bare_ink.get(weight) is None:
                    print(f"  {weight:<7} {kind:<9} no picture")
                    continue
                held.save(SCRATCH / f"{book_path.stem}.png")
                found = measure(bare_ink[weight], ink_of(held))
                if found is None:
                    print(f"  {weight:<7} {kind:<9} no ink")
                    continue
                named = size[0] if size else "default"
                rule_px = weight * 96 / 72
                print(f"  {weight:<7} {kind:<9} {named:<8} {found['length']:>11}"
                      f" {found['across']:>7} {found['past the end']:>13}"
                      f"      {found['length'] / rule_px:5.2f} long x {found['across'] / rule_px:5.2f} across")
            finally:
                book.Close(SaveChanges=False)
    finally:
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
