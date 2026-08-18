# -*- coding: utf-8 -*-
"""Export the line-height probe and compare the pitch with the font metrics."""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\lineheight\lineheight.pptx").resolve()
DST = SRC.with_suffix(".pdf")
FONTS = ["Arial", "Calibri", "Times New Roman", "Verdana", "Georgia",
         "Tahoma", "Segoe UI", "Trebuchet MS"]
SIZE = 40.0
EM = 2048

gdi = ctypes.WinDLL("gdi32.dll")
user = ctypes.WinDLL("user32.dll")
gdi.CreateFontW.restype = ctypes.c_void_p


class PANOSE(ctypes.Structure):
    _fields_ = [("b%d" % i, ctypes.c_ubyte) for i in range(10)]


class POINTL(ctypes.Structure):
    _fields_ = [("x", wt.LONG), ("y", wt.LONG)]


class RECTL(ctypes.Structure):
    _fields_ = [("left", wt.LONG), ("top", wt.LONG), ("right", wt.LONG), ("bottom", wt.LONG)]


class TEXTMETRICW(ctypes.Structure):
    _fields_ = [("tmHeight", wt.LONG), ("tmAscent", wt.LONG), ("tmDescent", wt.LONG),
                ("tmInternalLeading", wt.LONG), ("tmExternalLeading", wt.LONG),
                ("tmAveCharWidth", wt.LONG), ("tmMaxCharWidth", wt.LONG), ("tmWeight", wt.LONG),
                ("tmOverhang", wt.LONG), ("tmDigitizedAspectX", wt.LONG),
                ("tmDigitizedAspectY", wt.LONG), ("tmFirstChar", ctypes.c_wchar),
                ("tmLastChar", ctypes.c_wchar), ("tmDefaultChar", ctypes.c_wchar),
                ("tmBreakChar", ctypes.c_wchar), ("tmItalic", ctypes.c_ubyte),
                ("tmUnderlined", ctypes.c_ubyte), ("tmStruckOut", ctypes.c_ubyte),
                ("tmPitchAndFamily", ctypes.c_ubyte), ("tmCharSet", ctypes.c_ubyte)]


class OTM(ctypes.Structure):
    _fields_ = [("otmSize", wt.UINT), ("otmTextMetrics", TEXTMETRICW), ("otmFiller", ctypes.c_ubyte),
                ("otmPanoseNumber", PANOSE), ("otmfsSelection", wt.UINT), ("otmfsType", wt.UINT),
                ("otmsCharSlopeRise", ctypes.c_int), ("otmsCharSlopeRun", ctypes.c_int),
                ("otmItalicAngle", ctypes.c_int), ("otmEMSquare", wt.UINT), ("otmAscent", ctypes.c_int),
                ("otmDescent", ctypes.c_int), ("otmLineGap", wt.UINT), ("otmsCapEmHeight", wt.UINT),
                ("otmsXHeight", wt.UINT), ("otmrcFontBox", RECTL), ("otmMacAscent", ctypes.c_int),
                ("otmMacDescent", ctypes.c_int), ("otmMacLineGap", wt.UINT), ("otmusMinimumPPEM", wt.UINT),
                ("otmptSubscriptSize", POINTL), ("otmptSubscriptOffset", POINTL),
                ("otmptSuperscriptSize", POINTL), ("otmptSuperscriptOffset", POINTL),
                ("otmsStrikeoutSize", wt.UINT), ("otmsStrikeoutPosition", ctypes.c_int),
                ("otmsUnderscoreSize", ctypes.c_int), ("otmsUnderscorePosition", ctypes.c_int),
                ("otmpFamilyName", ctypes.c_char_p), ("otmpFaceName", ctypes.c_char_p),
                ("otmpStyleName", ctypes.c_char_p), ("otmpFullName", ctypes.c_char_p)]


def metrics(family: str) -> dict:
    hdc = ctypes.c_void_p(user.GetDC(None))
    hf = ctypes.c_void_p(gdi.CreateFontW(-EM, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, family))
    old = gdi.SelectObject(hdc, hf)
    size = gdi.GetOutlineTextMetricsW(hdc, 0, None)
    buf = ctypes.create_string_buffer(size)
    gdi.GetOutlineTextMetricsW(hdc, size, buf)
    o = ctypes.cast(buf, ctypes.POINTER(OTM)).contents
    tm = o.otmTextMetrics
    out = {
        "typo": (o.otmAscent - o.otmDescent + o.otmLineGap) / o.otmEMSquare,
        "hhea": (o.otmMacAscent - o.otmMacDescent + o.otmMacLineGap) / o.otmEMSquare,
        "tm": (tm.tmHeight + tm.tmExternalLeading) / o.otmEMSquare,
    }
    gdi.SelectObject(hdc, old)
    gdi.DeleteObject(hf)
    user.ReleaseDC(None, hdc)
    return out


def export() -> None:
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        prs = app.Presentations.Open(str(SRC), WithWindow=False)
        try:
            prs.SaveAs(str(DST), 32)
        finally:
            prs.Close()
    finally:
        app.Quit()
    print("exported", DST, DST.stat().st_size, "bytes")


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    doc = pymupdf.open(DST)
    print(f"{'font':>16} {'pitch':>7} {'ratio':>7} {'typo':>7} {'hhea':>7} {'tm':>7}  best")
    for i, fam in enumerate(FONTS):
        ys = []
        for b in doc[i].get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                cs = [c for s in l["spans"] for c in s["chars"]]
                if not cs or l["spans"][0]["size"] < SIZE - 2:
                    continue
                ys.append(cs[0]["origin"][1])
        ys.sort()
        if len(ys) < 2:
            print(f"{fam:>16}  (no lines)")
            continue
        pitch = (ys[-1] - ys[0]) / (len(ys) - 1)
        r = pitch / SIZE
        m = metrics(fam)
        best = min(m, key=lambda k: abs(m[k] - r))
        print(f"{fam:>16} {pitch:7.2f} {r:7.4f} {m['typo']:7.4f} {m['hhea']:7.4f} {m['tm']:7.4f}  {best}")
    doc.close()


if __name__ == "__main__":
    main()
