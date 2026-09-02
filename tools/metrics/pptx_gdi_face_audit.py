# -*- coding: utf-8 -*-
"""Which of this machine's faces does GDI measure differently from their files?

`gdi_font_view_can_be_corrupt` found one: asked for Caladea, GDI returns
advances 3.35% narrower than Caladea-Regular.ttf carries, per glyph and in both
directions, while Arial, Calibri and Georgia agree with their files to +0.00%.
PowerPoint reads the file (DirectWrite), Oxi asks GDI, and deck 47 breaks two
paragraphs early because of it.

That was found one family at a time, after seven hypotheses died. This asks the
whole question at once: for every family the engine may be asked to measure,
compare GDI's own text extent against the `hmtx` sum of the file the registry
points at. A family that disagrees is one where the engine's break test is
reading a lie -- and the fix for those is the design table, not another
rendering hypothesis.

★It needs THREE controls before any single number means anything: a family that
matches its file at +0.00% is what turns a -3.35% into a fact about the machine
rather than a fact about the instrument.

★This machine, 2026-09-02 (64 characters at a 100px em):

    Caladea            -3.36%   ← the one
    Calibri, Carlito   +0.51%   (Carlito is metric-compatible with Calibri and
    Arial, Lib Sans    +0.28%    reports the identical extent, as it should)
    Lib Serif, Times   -0.13%
    Georgia, Verdana   -0.06%
    Tahoma, Segoe UI   -0.02%

The +0.5% band is `GetTextExtentPoint32W` rounding each glyph to a whole pixel,
which is why it is positive for the narrow faces and near zero for the wide
ones. Caladea is six times that, negative, and per-glyph it scatters both ways
-- a different set of outlines, not a rounding.

    python tools/metrics/pptx_gdi_face_audit.py
    python tools/metrics/pptx_gdi_face_audit.py Caladea Carlito
    python tools/metrics/pptx_gdi_face_audit.py --corpus   # every family the decks name
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import sys
import winreg
from pathlib import Path

from fontTools.ttLib import TTFont, TTCollection

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

gdi = ctypes.WinDLL("gdi32")
user = ctypes.WinDLL("user32")

SAMPLE = ("The quick brown fox jumps over the lazy dog "
          "0123456789 ,.;:!?()-")
SIZE = 100  # a big em so a per-mille difference is visible in integer pixels

# Families to ask about when none are named: the three that have always agreed
# (the controls), then everything the local metric tables carry.
DEFAULT = [
    "Arial", "Calibri", "Georgia",
    "Caladea", "Carlito",
    "Liberation Sans", "Liberation Serif", "Liberation Mono",
    "Cambria", "Times New Roman", "Verdana", "Tahoma", "Segoe UI",
]


class SIZE_T(ctypes.Structure):
    _fields_ = [("cx", wt.LONG), ("cy", wt.LONG)]


def gdi_extent(family: str, size: int, text: str) -> int | None:
    """GDI's own width for `text`, in pixels, at `size` px em."""
    dc = user.GetDC(None)
    try:
        font = gdi.CreateFontW(-size, 0, 0, 0, 400, 0, 0, 0,
                               1, 0, 0, 0, 0, family)
        if not font:
            return None
        old = gdi.SelectObject(dc, font)
        # The face GDI actually served: asked for a name nothing has, the font
        # mapper answers with something else and this comparison would be
        # measuring the substitute.
        buf = ctypes.create_unicode_buffer(64)
        gdi.GetTextFaceW(dc, 64, buf)
        served = buf.value
        sz = SIZE_T()
        ok = gdi.GetTextExtentPoint32W(dc, text, len(text), ctypes.byref(sz))
        gdi.SelectObject(dc, old)
        gdi.DeleteObject(font)
        if not ok or not served.lower().startswith(family.lower()[:8]):
            return None
        return sz.cx
    finally:
        user.ReleaseDC(None, dc)


def registry_file(family: str) -> Path | None:
    """The file the machine says is this family's regular face."""
    for hive, key in ((winreg.HKEY_LOCAL_MACHINE,
                       r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
                      (winreg.HKEY_CURRENT_USER,
                       r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")):
        try:
            k = winreg.OpenKey(hive, key)
        except OSError:
            continue
        i = 0
        while True:
            try:
                name, value, _ = winreg.EnumValue(k, i)
            except OSError:
                break
            i += 1
            base = name.split("(")[0].strip()
            if base.lower() == family.lower():
                p = Path(value)
                if not p.is_absolute():
                    p = Path(r"C:\Windows\Fonts") / value
                if p.exists():
                    return p
    return None


def file_width(path: Path, text: str, size: int) -> float | None:
    """The `hmtx` sum for `text` at `size`, straight out of the file."""
    try:
        font = TTCollection(str(path))[0] if path.suffix.lower() == ".ttc" else TTFont(str(path))
    except Exception:
        return None
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    total = 0
    for ch in text:
        g = cmap.get(ord(ch))
        if g is None:
            return None
        total += hmtx[g][0]
    return total / upm * size


def corpus_families() -> list[str]:
    """Every typeface the decks name, so the claim can be corpus-wide.

    A family nothing on the machine serves drops out of the comparison on its
    own (`gdi_extent` returns None when the mapper substitutes), which is the
    right answer here: the engine measures those through the deck's embedded
    part, not through an installed file.
    """
    import re
    import zipfile
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from pptx_dump_ab import deck_paths

    seen: set[str] = set()
    for _, src in deck_paths("all"):
        try:
            with zipfile.ZipFile(src) as z:
                for name in z.namelist():
                    if re.match(r"ppt/(slides|slideLayouts|slideMasters|theme)/[^/]+\.xml$", name):
                        xml = z.read(name).decode("utf-8", "replace")
                        seen |= {t for t in re.findall(r'typeface="([^"]*)"', xml)
                                 if t and not t.startswith("+")}
        except Exception:
            continue
    return sorted(seen)


def main() -> None:
    argv = sys.argv[1:]
    if argv and argv[0] == "--corpus":
        families = corpus_families()
        print(f"{len(families)} families named by the corpus\n")
    else:
        families = argv or DEFAULT
    print(f"{len(SAMPLE)} characters at {SIZE}px em\n")
    print(f"{'family':<22}{'GDI':>10}{'file':>12}{'delta':>10}   file")
    rows: list[tuple[float, str]] = []
    for fam in families:
        got = gdi_extent(fam, SIZE, SAMPLE)
        path = registry_file(fam)
        want = file_width(path, SAMPLE, SIZE) if path else None
        if got is None:
            print(f"{fam:<22}{'not served':>10}")
            continue
        if want is None:
            print(f"{fam:<22}{got:>10}{'no file':>12}")
            continue
        delta = (got - want) / want * 100.0
        rows.append((abs(delta), fam))
        print(f"{fam:<22}{got:>10}{want:>12.1f}{delta:>9.2f}%   {path.name}")
    # ★The denominator, so a short list of hits is never read as a clean sweep:
    # most families a deck names are not on the machine at all -- they arrive as
    # embedded parts -- and those are simply outside this question.
    if rows:
        rows.sort(reverse=True)
        band = [d for d, _ in rows]
        print(f"\n{len(families)} families asked, {len(rows)} served by this machine "
              f"and comparable, {len(families) - len(rows)} not served or with no file")
        bad = [(d, f) for d, f in rows if d > 1.5]
        print(f"{len(bad)} disagree with their file by more than 1.5%"
              + (": " + ", ".join(f"{f} {d:.2f}%" for d, f in bad) if bad else ""))
        print(f"the rest sit within {max(d for d, _ in rows if d <= 1.5):.2f}%, "
              f"which is the per-glyph pixel rounding of the extent call"
              if any(d <= 1.5 for d in band) else "")


if __name__ == "__main__":
    main()
