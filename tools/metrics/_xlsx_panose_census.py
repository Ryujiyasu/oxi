# -*- coding: utf-8 -*-
r"""Which faces does the corpus ask for that this machine does not have, and
what PANOSE do they carry?

SX47 found that Excel replaces a missing face by matching its PANOSE against
the installed ones — `AR P丸ゴシック体E` with its PANOSE draws as 游ゴシック to
the pixel, and GDI, which has nowhere to put a PANOSE, hands back ＭＳ Ｐゴシック
instead. Writing that matcher needs two things this prints: the missing faces
worth matching, and what every installed face's PANOSE is.

    python tools\metrics\_xlsx_panose_census.py
    python tools\metrics\_xlsx_panose_census.py --installed
"""
import argparse
import ctypes
import re
import sys
import zipfile
from collections import Counter, defaultdict
from ctypes import wintypes
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
DOCS = REPO / "tools" / "golden-test" / "documents" / "xlsx"

gdi = ctypes.windll.gdi32
user = ctypes.windll.user32


class LOGFONTW(ctypes.Structure):
    _fields_ = [("lfHeight", wintypes.LONG), ("lfWidth", wintypes.LONG),
                ("lfEscapement", wintypes.LONG), ("lfOrientation", wintypes.LONG),
                ("lfWeight", wintypes.LONG), ("lfItalic", ctypes.c_byte),
                ("lfUnderline", ctypes.c_byte), ("lfStrikeOut", ctypes.c_byte),
                ("lfCharSet", ctypes.c_byte), ("lfOutPrecision", ctypes.c_byte),
                ("lfClipPrecision", ctypes.c_byte), ("lfQuality", ctypes.c_byte),
                ("lfPitchAndFamily", ctypes.c_byte),
                ("lfFaceName", wintypes.WCHAR * 32)]


class PANOSE(ctypes.Structure):
    _fields_ = [(name, ctypes.c_byte) for name in
                ("bFamilyType", "bSerifStyle", "bWeight", "bProportion",
                 "bContrast", "bStrokeVariation", "bArmStyle", "bLetterform",
                 "bMidline", "bXHeight")]


class TEXTMETRICW(ctypes.Structure):
    _fields_ = [("tmHeight", wintypes.LONG), ("tmAscent", wintypes.LONG),
                ("tmDescent", wintypes.LONG), ("tmInternalLeading", wintypes.LONG),
                ("tmExternalLeading", wintypes.LONG), ("tmAveCharWidth", wintypes.LONG),
                ("tmMaxCharWidth", wintypes.LONG), ("tmWeight", wintypes.LONG),
                ("tmOverhang", wintypes.LONG), ("tmDigitizedAspectX", wintypes.LONG),
                ("tmDigitizedAspectY", wintypes.LONG), ("tmFirstChar", wintypes.WCHAR),
                ("tmLastChar", wintypes.WCHAR), ("tmDefaultChar", wintypes.WCHAR),
                ("tmBreakChar", wintypes.WCHAR), ("tmItalic", ctypes.c_byte),
                ("tmUnderlined", ctypes.c_byte), ("tmStruckOut", ctypes.c_byte),
                ("tmPitchAndFamily", ctypes.c_byte), ("tmCharSet", ctypes.c_byte)]


class OUTLINETEXTMETRICW(ctypes.Structure):
    _fields_ = [("otmSize", wintypes.UINT), ("otmTextMetrics", TEXTMETRICW),
                ("otmFiller", ctypes.c_byte), ("otmPanoseNumber", PANOSE),
                ("otmfsSelection", wintypes.UINT), ("otmfsType", wintypes.UINT),
                ("otmsCharSlopeRise", wintypes.LONG), ("otmsCharSlopeRun", wintypes.LONG),
                ("otmItalicAngle", wintypes.LONG), ("otmEMSquare", wintypes.UINT),
                ("otmAscent", wintypes.LONG), ("otmDescent", wintypes.LONG),
                ("otmLineGap", wintypes.UINT), ("otmsCapEmHeight", wintypes.UINT),
                ("otmsXHeight", wintypes.UINT),
                ("otmrcFontBox", wintypes.RECT), ("otmMacAscent", wintypes.LONG),
                ("otmMacDescent", wintypes.LONG), ("otmMacLineGap", wintypes.UINT),
                ("otmusMinimumPPEM", wintypes.UINT),
                ("otmptSubscriptSize", wintypes.POINT),
                ("otmptSubscriptOffset", wintypes.POINT),
                ("otmptSuperscriptSize", wintypes.POINT),
                ("otmptSuperscriptOffset", wintypes.POINT),
                ("otmsStrikeoutSize", wintypes.UINT),
                ("otmsStrikeoutPosition", wintypes.LONG),
                ("otmsUnderscoreSize", wintypes.LONG),
                ("otmsUnderscorePosition", wintypes.LONG),
                ("otmpFamilyName", wintypes.LPARAM), ("otmpFaceName", wintypes.LPARAM),
                ("otmpStyleName", wintypes.LPARAM), ("otmpFullName", wintypes.LPARAM)]


ENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.POINTER(LOGFONTW),
                              ctypes.c_void_p, wintypes.DWORD, wintypes.LPARAM)


def installed():
    """Every face this machine has, by name, with its charset."""
    found = {}

    def take(logfont, _metric, _kind, _held):
        face = logfont.contents.lfFaceName
        if face and not face.startswith("@"):
            found.setdefault(face, logfont.contents.lfCharSet)
        return 1

    dc = user.GetDC(0)
    wanted = LOGFONTW()
    wanted.lfCharSet = 1                     # DEFAULT_CHARSET: every script
    gdi.EnumFontFamiliesExW(dc, ctypes.byref(wanted), ENUMPROC(take), 0, 0)
    user.ReleaseDC(0, dc)
    return found


def panose_of(face):
    """The face's own PANOSE, and the name GDI answers with."""
    dc = user.GetDC(0)
    memory = gdi.CreateCompatibleDC(dc)
    font = gdi.CreateFontW(-32, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, face)
    gdi.SelectObject(memory, font)
    size = gdi.GetOutlineTextMetricsW(memory, 0, None)
    answer, panose = None, None
    if size:
        held = ctypes.create_string_buffer(size)
        gdi.GetOutlineTextMetricsW(memory, size, held)
        metric = ctypes.cast(held, ctypes.POINTER(OUTLINETEXTMETRICW)).contents
        panose = tuple(getattr(metric.otmPanoseNumber, name) & 0xFF
                       for name, _ in PANOSE._fields_)
        at = metric.otmpFaceName
        answer = ctypes.wstring_at(ctypes.addressof(held) + at) if at else None
    gdi.DeleteObject(font)
    gdi.DeleteDC(memory)
    user.ReleaseDC(0, dc)
    return panose, answer


def asked_for():
    """Every face the corpus names, and the PANOSE the file states for it."""
    faces = Counter()
    panoses = defaultdict(Counter)
    where = defaultdict(set)
    for book in sorted(DOCS.glob("*.xlsx")):
        try:
            held = zipfile.ZipFile(book)
        except Exception:
            continue
        for part in held.namelist():
            if not (part.endswith(".xml") and
                    ("styles" in part or "drawings/" in part or "charts/" in part)):
                continue
            try:
                xml = held.read(part).decode("utf-8", "replace")
            except Exception:
                continue
            # A cell font: <name val="…"/>. A drawing run: <a:latin
            # typeface="…" panose="…"/>, and its ea and cs siblings.
            for name in re.findall(r'<name val="([^"]+)"', xml):
                faces[name] += 1
                where[name].add(book.stem)
            for tag, attrs in re.findall(r"<a:(latin|ea|cs)\s+([^/>]*)/?>", xml):
                face = re.search(r'typeface="([^"]*)"', attrs)
                if not face or not face.group(1):
                    continue
                faces[face.group(1)] += 1
                where[face.group(1)].add(book.stem)
                panose = re.search(r'panose="([0-9A-Fa-f]+)"', attrs)
                if panose:
                    panoses[face.group(1)][panose.group(1)] += 1
    return faces, panoses, where


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--installed", action="store_true",
                        help="print every installed face and its PANOSE")
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    have = installed()
    if args.installed:
        print(f"{'face':<34}{'charset':>8}  PANOSE")
        for face in sorted(have):
            panose, _ = panose_of(face)
            if panose:
                print(f"{face:<34}{have[face]:>8}  "
                      + " ".join(f"{number:>3}" for number in panose))
        return

    faces, panoses, where = asked_for()
    missing = [(name, count) for name, count in faces.most_common()
               if name not in have]
    print(f"{len(faces)} faces asked for, {len(missing)} of them not installed\n")
    print(f"{'face':<28}{'uses':>6}{'books':>7}  {'stated PANOSE':<26}GDI answers")
    for name, count in missing:
        stated = ", ".join(panose for panose, _ in panoses[name].most_common(2))
        _, answer = panose_of(name)
        print(f"{name:<28}{count:>6}{len(where[name]):>7}  {stated or '-':<26}{answer}")


if __name__ == "__main__":
    main()
