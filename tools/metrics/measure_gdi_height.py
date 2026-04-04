"""Measure GDI tmHeight/tmAscent/tmDescent for fonts missing from gdi_height_table.json.

Uses Win32 GetTextMetrics via ctypes to get exact GDI pixel values.
"""
import ctypes
from ctypes import wintypes
import json
import os

# Win32 GDI constants
LF_FACESIZE = 32
CLEARTYPE_QUALITY = 5
DEFAULT_CHARSET = 1

class TEXTMETRICW(ctypes.Structure):
    _fields_ = [
        ("tmHeight", wintypes.LONG),
        ("tmAscent", wintypes.LONG),
        ("tmDescent", wintypes.LONG),
        ("tmInternalLeading", wintypes.LONG),
        ("tmExternalLeading", wintypes.LONG),
        ("tmAveCharWidth", wintypes.LONG),
        ("tmMaxCharWidth", wintypes.LONG),
        ("tmWeight", wintypes.LONG),
        ("tmOverhang", wintypes.LONG),
        ("tmDigitizedAspectX", wintypes.LONG),
        ("tmDigitizedAspectY", wintypes.LONG),
        ("tmFirstChar", wintypes.WORD),
        ("tmLastChar", wintypes.WORD),
        ("tmDefaultChar", wintypes.WORD),
        ("tmBreakChar", wintypes.WORD),
        ("tmItalic", ctypes.c_byte),
        ("tmUnderlined", ctypes.c_byte),
        ("tmStruckOut", ctypes.c_byte),
        ("tmPitchAndFamily", ctypes.c_byte),
        ("tmCharSet", ctypes.c_byte),
    ]

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

CreateFontW = gdi32.CreateFontW
CreateFontW.restype = ctypes.c_void_p
SelectObject = gdi32.SelectObject
SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
SelectObject.restype = ctypes.c_void_p
DeleteObject = gdi32.DeleteObject
DeleteObject.argtypes = [ctypes.c_void_p]
GetTextMetricsW = gdi32.GetTextMetricsW
GetTextMetricsW.argtypes = [ctypes.c_void_p, ctypes.POINTER(TEXTMETRICW)]
GetDC = user32.GetDC
GetDC.restype = ctypes.c_void_p
ReleaseDC = user32.ReleaseDC
CreateCompatibleDC = gdi32.CreateCompatibleDC
CreateCompatibleDC.argtypes = [ctypes.c_void_p]
CreateCompatibleDC.restype = ctypes.c_void_p
DeleteDC = gdi32.DeleteDC
DeleteDC.argtypes = [ctypes.c_void_p]


def measure_font(font_name: str, bold: bool = False, ppem_range=(5, 101)):
    """Measure tmHeight, tmAscent, tmDescent for each ppem."""
    dc = CreateCompatibleDC(0)
    results = {}
    weight = 700 if bold else 400

    for ppem in range(ppem_range[0], ppem_range[1]):
        font_name_buf = ctypes.create_unicode_buffer(font_name)
        hfont = CreateFontW(
            -ppem,  # negative = character height (ppem)
            0, 0, 0, weight,
            0, 0, 0,
            DEFAULT_CHARSET,
            0, 0,
            CLEARTYPE_QUALITY,
            0,
            font_name_buf,
        )
        old = SelectObject(dc, hfont)
        tm = TEXTMETRICW()
        GetTextMetricsW(dc, ctypes.byref(tm))
        results[str(ppem)] = {
            "h": tm.tmHeight,
            "a": tm.tmAscent,
            "d": tm.tmDescent,
        }
        SelectObject(dc, old)
        DeleteObject(hfont)

    DeleteDC(dc)
    return results


def main():
    # Fonts to measure (missing from current table)
    fonts_to_measure = [
        ("Cambria", True, "Cambria_Bold"),
        ("Century", True, "Century_Bold"),
        ("Yu Mincho Demibold", False, "Yu_Mincho_Demibold"),
        ("HGGothicE", False, "HGGothicE"),
        ("HGMinchoE", False, "HGMinchoE"),
        ("HGSGothicE", False, "HGSGothicE"),
        ("HGSMinchoE", False, "HGSMinchoE"),
    ]

    # Load existing table
    table_path = os.path.join(os.path.dirname(__file__), "..", "..",
                              "crates", "oxidocs-core", "src", "font", "data",
                              "gdi_height_table.json")
    with open(table_path, encoding="utf-8") as f:
        table = json.load(f)

    for font_name, bold, key in fonts_to_measure:
        if key in table:
            print(f"[SKIP] {key} already in table")
            continue
        print(f"Measuring {key} ({font_name}, bold={bold})...")
        data = measure_font(font_name, bold=bold)
        # Verify: check a few values
        p15 = data.get("15", {})
        p20 = data.get("20", {})
        print(f"  ppem=15: h={p15.get('h')} a={p15.get('a')} d={p15.get('d')}")
        print(f"  ppem=20: h={p20.get('h')} a={p20.get('a')} d={p20.get('d')}")
        # Only add if h > 0 (font exists)
        if p15.get("h", 0) > 0:
            table[key] = data
            print(f"  -> Added {len(data)} entries")
        else:
            print(f"  -> SKIPPED (font not found or h=0)")

    # Save
    with open(table_path, "w", encoding="utf-8") as f:
        json.dump(table, f, separators=(",", ":"))
    print(f"\nSaved to {table_path}")
    print(f"Total fonts: {len(table)}")


if __name__ == "__main__":
    main()
