"""
§1.7 follow-up: per-font ascent/descent via GDI TEXTMETRIC
to test the "mixed_line_h = max(asc) + max(desc)" hypothesis.

Word COM doesn't expose Font.Ascent/Descent directly. Use Win32 GDI:
  CreateFontW + GetTextMetricsW → tmAscent, tmDescent (in pixels).

Convert to pt: pt = px * 72 / 96 (assuming 96 dpi).

For each font/size combination, capture (ascent_pt, descent_pt, total_pt).
Then compare mixed_line_h vs max-ascent + max-descent across pairs.
"""
import ctypes
from ctypes import wintypes
import os
import sys
import json

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_JSON = os.path.join(os.path.dirname(__file__), "output",
                        "ra2_font_asc_desc_metrics.json")

# Win32 setup
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
user32 = ctypes.WinDLL("user32", use_last_error=True)

# CreateFontW signature
gdi32.CreateFontW.argtypes = [
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    ctypes.c_int, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint,
    ctypes.c_uint, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint,
    ctypes.c_uint, ctypes.c_wchar_p,
]
gdi32.CreateFontW.restype = wintypes.HFONT
gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
gdi32.SelectObject.restype = wintypes.HGDIOBJ
gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
gdi32.DeleteObject.restype = wintypes.BOOL
gdi32.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
gdi32.GetDeviceCaps.restype = ctypes.c_int


class TEXTMETRICW(ctypes.Structure):
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


gdi32.GetTextMetricsW.argtypes = [wintypes.HDC, ctypes.POINTER(TEXTMETRICW)]
gdi32.GetTextMetricsW.restype = wintypes.BOOL


def measure_font_metrics(family, size_pt):
    """Return dict with ascent/descent/height in pt (96dpi)."""
    hdc = user32.GetDC(0)
    try:
        # Create font: -size means request character height in logical units (px at 96dpi)
        size_px = round(size_pt * 96.0 / 72.0)
        hfont = gdi32.CreateFontW(
            -size_px, 0, 0, 0,
            400, 0, 0, 0,  # normal weight, no italic/underline/strikeout
            1,  # DEFAULT_CHARSET
            0, 0,
            5,  # CLEARTYPE_QUALITY
            0,
            family,
        )
        old_font = gdi32.SelectObject(hdc, hfont)
        tm = TEXTMETRICW()
        ok = gdi32.GetTextMetricsW(hdc, ctypes.byref(tm))
        if not ok:
            return None
        gdi32.SelectObject(hdc, old_font)
        gdi32.DeleteObject(hfont)
        # Convert px → pt
        return {
            "family": family,
            "size_pt": size_pt,
            "ppem": size_px,
            "tmAscent_px": tm.tmAscent,
            "tmDescent_px": tm.tmDescent,
            "tmHeight_px": tm.tmHeight,
            "tmInternalLeading_px": tm.tmInternalLeading,
            "ascent_pt": round(tm.tmAscent * 72.0 / 96.0, 4),
            "descent_pt": round(tm.tmDescent * 72.0 / 96.0, 4),
            "height_pt": round(tm.tmHeight * 72.0 / 96.0, 4),
        }
    finally:
        user32.ReleaseDC(0, hdc)


def main():
    fonts = ["Calibri", "Times New Roman", "MS Mincho", "MS Gothic",
             "Yu Mincho", "Yu Gothic"]
    sizes = [8, 10.5, 11, 14, 18, 24]
    results = []
    print(f"{'family':22s}{'size':>5}  {'asc':>6}  {'desc':>6}  {'height':>7}  {'iLead':>6}")
    print("-" * 60)
    for f in fonts:
        for s in sizes:
            m = measure_font_metrics(f, s)
            if m:
                results.append(m)
                print(f"{m['family']:22s}{s:>5}  "
                      f"{m['ascent_pt']:>6}  {m['descent_pt']:>6}  "
                      f"{m['height_pt']:>7}  {m['tmInternalLeading_px']*72/96:>6.2f}")

    # Cross-check: predict mixed_line_h via max-asc + max-desc for the v2 sweep
    print("\n=== Hypothesis test: mixed_line_h = max-asc + max-desc ===")
    # Pairs from v2 measurement
    v2_pairs = [
        ("Calibri", 11, "MS Mincho", 14, 28.5, 25.0, 35.5, "grid18"),
        ("Calibri", 14, "MS Mincho", 11, 18.0, 19.5, 16.0, "grid18"),
        ("Calibri", 18, "MS Mincho", 24, 40.0, 31.0, 36.0, "grid18"),
        ("Calibri", 8,  "MS Mincho", 8,  18.5, 17.5, 17.5, "grid18"),
        ("Times New Roman", 11, "MS Gothic", 14, 29.0, 24.5, 35.5, "grid18"),
        ("Times New Roman", 18, "MS Gothic", 24, 40.5, 30.5, 36.0, "grid18"),
        ("Yu Mincho", 8, "Yu Gothic", 8, 18.0, 18.0, 17.5, "grid18"),
        ("Yu Mincho", 11, "Yu Gothic", 14, 38.0, 33.0, 36.0, "grid18"),
    ]
    print(f"{'A':16s}{'sA':>4} / {'B':16s}{'sB':>4}  "
          f"{'mix':>6}  {'gapA':>6}  {'gapB':>6}  "
          f"{'maxAsc+maxDesc':>14}  {'(grid_snap)':>12}  {'pred':>6}  {'err':>6}")
    print("-" * 110)
    by_key = {(r["family"], r["size_pt"]): r for r in results}
    for fa, sa, fb, sb, mix, gA, gB, glabel in v2_pairs:
        ma = by_key.get((fa, sa))
        mb = by_key.get((fb, sb))
        if not (ma and mb):
            continue
        max_asc = max(ma["ascent_pt"], mb["ascent_pt"])
        max_desc = max(ma["descent_pt"], mb["descent_pt"])
        natural = max_asc + max_desc
        # Grid snap to next 18pt multiple (grid pitch 18pt)
        if glabel == "grid18":
            cells = max(1, (natural + 17.99) // 18)  # ceil(natural / 18)
            pred = cells * 18.0
        else:
            pred = natural
        err = mix - pred
        print(f"{fa:16s}{sa:>4} / {fb:16s}{sb:>4}  "
              f"{mix:>6}  {gA:>6}  {gB:>6}  "
              f"{natural:>14.2f}  {pred:>12.1f}  {pred:>6}  {err:+.2f}")

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved {len(results)} font metric records to {OUT_JSON}")


if __name__ == "__main__":
    main()
