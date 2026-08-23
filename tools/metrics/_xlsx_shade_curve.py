"""What does `shade` do to a colour — halve the byte, or halve the light?

`cas-r02-ippan-4hyou` rules a box in `schemeClr lt1` under `shade 50000`.
Excel draws 188,188,188. Halving the sRGB byte gives 128, which is what Oxi
draws. Halving the LIGHT and re-encoding gives 188, because sRGB is not
linear. One book is not enough to tell a curve from a coincidence, so this
sweeps the shade and the tint over several base colours and reads the drawn
byte off Excel's own picture.

The shapes are written into the drawing by hand — the modifiers are not
reachable through COM — and Excel is asked to draw the file.

Run: python tools/metrics/_xlsx_shade_curve.py
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

SCRATCH = Path(r"C:\tmp\xlsx_shade")
BASES = ["FFFFFF", "4472C4", "FF0000", "808080"]
AMOUNTS = [25000, 50000, 75000]
TOP_PT, LEFT_PT = 30.0, 30.0
SIDE_PT = 40.0


def seed() -> Path:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    made = SCRATCH / "seed.xlsx"
    if made.exists():
        return made
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:Z40").Interior.Color = 0xFFFFFF
        for at in range(len(BASES) * len(AMOUNTS) * 2):
            shape = sheet.Shapes.AddShape(
                1, LEFT_PT + (at % 8) * (SIDE_PT + 8), TOP_PT + (at // 8) * (SIDE_PT + 8),
                SIDE_PT, SIDE_PT,
            )
            shape.Line.Visible = False
        book.SaveAs(str(made), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return made


def written(source: Path, arms: list[tuple[str, str, int]]) -> Path:
    out = SCRATCH / "painted.xlsx"
    if out.exists():
        out.unlink()
    held = zipfile.ZipFile(source)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as writing:
        for item in held.infolist():
            raw = held.read(item.filename)
            if item.filename.startswith("xl/drawings/") and item.filename.endswith(".xml"):
                text = raw.decode("utf-8")
                fills = list(re.finditer(r"<a:solidFill>.*?</a:solidFill>", text, re.S))
                assert len(fills) >= len(arms), f"{len(fills)} fills for {len(arms)} arms"
                for spot, (base, kind, amount) in zip(reversed(fills[: len(arms)]),
                                                      reversed(arms)):
                    fresh = (f'<a:solidFill><a:srgbClr val="{base}">'
                             f'<a:{kind} val="{amount}"/></a:srgbClr></a:solidFill>')
                    text = text[: spot.start()] + fresh + text[spot.end():]
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


def straight(byte: float) -> float:
    """sRGB byte to linear light."""
    held = byte / 255.0
    return held / 12.92 if held <= 0.04045 else ((held + 0.055) / 1.055) ** 2.4


def encoded(light: float) -> float:
    """Linear light back to an sRGB byte."""
    held = 12.92 * light if light <= 0.0031308 else 1.055 * light ** (1 / 2.4) - 0.055
    return held * 255.0


def main() -> int:
    source = seed()
    arms = [(base, kind, amount)
            for base in BASES for kind in ("shade", "tint") for amount in AMOUNTS]
    arms = arms[: len(BASES) * len(AMOUNTS) * 2]
    book_path = written(source, arms)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Open(str(book_path))
    try:
        held = picture(book.Worksheets(1))
        if held is None:
            print("Excel would not hand over a picture")
            return 1
        held.save(SCRATCH / "shot.png")
        pixels = np.asarray(held.convert("RGB")).astype(float)
        # Where each shape sits and what it was painted with, read back out of
        # the file itself — guessing the order the fills appear in put every
        # reading against the wrong arm.
        holds = zipfile.ZipFile(book_path)
        found = []
        for part in holds.namelist():
            if not (part.startswith("xl/drawings/") and part.endswith(".xml")):
                continue
            text = holds.read(part).decode("utf-8", "replace")
            for block in re.finditer(r"<xdr:sp[ >].*?</xdr:sp>", text, re.S):
                body = block.group(0)
                off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', body)
                ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', body)
                paint = re.search(
                    r'<a:srgbClr val="([0-9A-Fa-f]{6})"><a:(shade|tint) val="(\d+)"/>', body
                )
                if off and ext and paint:
                    found.append((
                        int(off.group(1)) / 9525, int(off.group(2)) / 9525,
                        int(ext.group(1)) / 9525, int(ext.group(2)) / 9525,
                        paint.group(1), paint.group(2), int(paint.group(3)),
                    ))
        print(f"  {len(found)} painted shapes read back")
        print("  base    what   amount    Excel   byte x k   light x k   which fits")
        for x0, y0, wide, high, base, kind, amount in found:
            x = round(x0 + wide / 2)
            y = round(y0 + high / 2)
            seen = pixels[y, x]
            channel = int(base[0:2], 16)
            share = amount / 100_000
            if kind == "shade":
                by_byte = channel * share
                by_light = encoded(straight(channel) * share)
            else:
                by_byte = channel + (255 - channel) * share
                by_light = encoded(straight(channel) + (1 - straight(channel)) * share)
            fits = ("light" if abs(seen[0] - by_light) < abs(seen[0] - by_byte)
                    else "byte" if abs(seen[0] - by_byte) < abs(seen[0] - by_light) else "tie")
            print(f"  {base}  {kind:<6} {amount:<8} {str([int(v) for v in seen]):<15}"
                  f"{by_byte:7.1f}    {by_light:7.1f}    {fits}")
    finally:
        book.Close(SaveChanges=False)
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
