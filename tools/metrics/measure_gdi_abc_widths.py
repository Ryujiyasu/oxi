"""Measure GDI GetCharABCWidths for fonts to find actual pixel widths.

Compares font file advance widths vs GDI ABC widths to detect font linking.
"""
import ctypes
from ctypes import wintypes
import json
import os

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

class ABC(ctypes.Structure):
    _fields_ = [
        ("abcA", ctypes.c_int),
        ("abcB", ctypes.c_uint),
        ("abcC", ctypes.c_int),
    ]

class TEXTMETRIC(ctypes.Structure):
    _fields_ = [
        ("tmHeight", ctypes.c_long),
        ("tmAscent", ctypes.c_long),
        ("tmDescent", ctypes.c_long),
        ("tmInternalLeading", ctypes.c_long),
        ("tmExternalLeading", ctypes.c_long),
        ("tmAveCharWidth", ctypes.c_long),
        ("tmMaxCharWidth", ctypes.c_long),
        ("tmWeight", ctypes.c_long),
        ("tmOverhang", ctypes.c_long),
        ("tmDigitizedAspectX", ctypes.c_long),
        ("tmDigitizedAspectY", ctypes.c_long),
        ("tmFirstChar", ctypes.c_wchar),
        ("tmLastChar", ctypes.c_wchar),
        ("tmDefaultChar", ctypes.c_wchar),
        ("tmBreakChar", ctypes.c_wchar),
        ("tmItalic", ctypes.c_byte),
        ("tmUnderlined", ctypes.c_byte),
        ("tmStruckOut", ctypes.c_byte),
        ("tmPitchAndFamily", ctypes.c_byte),
        ("tmCharSet", ctypes.c_byte),
    ]

def measure_font_widths(font_name, ppem, chars):
    """Get GDI character widths for a font at given ppem."""
    hdc = user32.GetDC(0)

    # Create font at specified ppem (negative height = character height)
    hfont = gdi32.CreateFontW(
        -ppem,  # nHeight (negative = character height in pixels)
        0,      # nWidth
        0, 0,   # escapement, orientation
        400,    # weight (normal)
        0, 0, 0,  # italic, underline, strikeout
        128,    # charset: SHIFTJIS_CHARSET for Japanese fonts
        0, 0, 0, 0,  # out precision, clip, quality, pitch
        font_name  # face name
    )

    old_font = gdi32.SelectObject(hdc, hfont)

    # Get TEXTMETRIC
    tm = TEXTMETRIC()
    gdi32.GetTextMetricsW(hdc, ctypes.byref(tm))

    results = {
        "font": font_name,
        "ppem": ppem,
        "tmHeight": tm.tmHeight,
        "tmAscent": tm.tmAscent,
        "tmDescent": tm.tmDescent,
        "tmAveCharWidth": tm.tmAveCharWidth,
        "tmMaxCharWidth": tm.tmMaxCharWidth,
        "tmOverhang": tm.tmOverhang,
        "chars": {}
    }

    for c in chars:
        cp = ord(c)
        abc = ABC()
        ok = gdi32.GetCharABCWidthsW(hdc, cp, cp, ctypes.byref(abc))
        if ok:
            total = abc.abcA + abc.abcB + abc.abcC
            results["chars"][c] = {
                "A": abc.abcA, "B": abc.abcB, "C": abc.abcC,
                "total": total,
                "pt": round(total * 72 / 96, 2)
            }
        else:
            # Fallback to GetCharWidth32
            width = ctypes.c_int()
            gdi32.GetCharWidth32W(hdc, cp, cp, ctypes.byref(width))
            results["chars"][c] = {
                "width32": width.value,
                "pt": round(width.value * 72 / 96, 2)
            }

    gdi32.SelectObject(hdc, old_font)
    gdi32.DeleteObject(hfont)
    user32.ReleaseDC(0, hdc)

    return results


if __name__ == "__main__":
    # Test characters
    latin_chars = list("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .,;:!?-()[]{}\"'/@#$%&*+=<>")
    cjk_chars = list("あいうえおかきくけこアイウエオカキクケコ漢字日本語")
    all_chars = latin_chars + cjk_chars

    # Fonts to test
    fonts = [
        ("ＭＳ 明朝", [14, 16, 20]),   # MS Mincho - CJK name
        ("MS 明朝", [14, 16, 20]),      # MS Mincho - alt name
        ("ＭＳ ゴシック", [14, 16, 20]), # MS Gothic
    ]

    output = {}
    for font_name, ppems in fonts:
        for ppem in ppems:
            key = f"{font_name}@{ppem}"
            print(f"Measuring {key}...")
            result = measure_font_widths(font_name, ppem, all_chars)
            output[key] = result

            # Print summary
            print(f"  tmHeight={result['tmHeight']} tmAve={result['tmAveCharWidth']} tmMax={result['tmMaxCharWidth']}")
            # Show a few Latin chars
            for c in ['T', 'h', 'i', 's', 'm', 'w']:
                if c in result['chars']:
                    d = result['chars'][c]
                    if 'A' in d:
                        print(f"  {c}: A={d['A']} B={d['B']} C={d['C']} total={d['total']}px = {d['pt']}pt")
                    else:
                        print(f"  {c}: width32={d['width32']}px = {d['pt']}pt")

    outpath = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data", "gdi_abc_widths.json")
    with open(outpath, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {outpath}")
