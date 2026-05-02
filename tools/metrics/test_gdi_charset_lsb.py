"""Investigation B (direct GDI test): does CreateFontW charset affect □ glyph LSB?

Hypothesis: Oxi's `tools/oxi-gdi-renderer/src/main.rs:193` uses DEFAULT_CHARSET (1).
With DEFAULT_CHARSET, GDI may resolve "MS Gothic" to a different physical font
than with SHIFTJIS_CHARSET (128), producing different glyph LSB for □ U+25A1.

Test:
  - For each charset ∈ {DEFAULT=1, SHIFTJIS=128, ANSI=0}:
    1. CreateFontW with that charset, family="MS Gothic", height=-19 (= 14pt @ 96dpi)
    2. Get TEXTMETRIC + glyph metrics via GetGlyphOutlineW or GetCharABCWidthsW
    3. Render □ via TextOutW at known x, capture pixel
    4. Find leftmost dark pixel in glyph row
    5. Report LSB = pixel_left - x

If charset matters: at least one charset produces different LSB than others.

If all same: hypothesis falsified, look elsewhere (lfQuality, lfPitchAndFamily,
or font registry mismatch).
"""
import ctypes
import ctypes.wintypes as wt
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Load libraries
gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

# Constants
DEFAULT_CHARSET = 1
SHIFTJIS_CHARSET = 128
ANSI_CHARSET = 0
SYMBOL_CHARSET = 2

OUT_DEFAULT_PRECIS = 0
CLIP_DEFAULT_PRECIS = 0
CLEARTYPE_QUALITY = 5
DEFAULT_PITCH = 0

CAP_HEIGHT_PT = 14.0
PPEM = round(CAP_HEIGHT_PT * 96 / 72)  # = 19

MM_TEXT = 1
TRANSPARENT = 1
WHITE_BRUSH = 0
BLACK_PEN = 7

# Bitmap size
BMP_W = 100
BMP_H = 50


class TEXTMETRICW(ctypes.Structure):
    _fields_ = [
        ("tmHeight", wt.LONG),
        ("tmAscent", wt.LONG),
        ("tmDescent", wt.LONG),
        ("tmInternalLeading", wt.LONG),
        ("tmExternalLeading", wt.LONG),
        ("tmAveCharWidth", wt.LONG),
        ("tmMaxCharWidth", wt.LONG),
        ("tmWeight", wt.LONG),
        ("tmOverhang", wt.LONG),
        ("tmDigitizedAspectX", wt.LONG),
        ("tmDigitizedAspectY", wt.LONG),
        ("tmFirstChar", wt.WCHAR),
        ("tmLastChar", wt.WCHAR),
        ("tmDefaultChar", wt.WCHAR),
        ("tmBreakChar", wt.WCHAR),
        ("tmItalic", wt.BYTE),
        ("tmUnderlined", wt.BYTE),
        ("tmStruckOut", wt.BYTE),
        ("tmPitchAndFamily", wt.BYTE),
        ("tmCharSet", wt.BYTE),
    ]


class ABC(ctypes.Structure):
    _fields_ = [("abcA", wt.LONG), ("abcB", wt.UINT), ("abcC", wt.LONG)]


def render_char_with_charset(charset: int, char: str, x: int = 20, y: int = 5) -> dict:
    """Create font with given charset, render `char` at (x, y), return pixel data + metrics."""
    # Create memory DC
    screen_dc = user32.GetDC(0)
    mem_dc = gdi32.CreateCompatibleDC(screen_dc)

    # Create bitmap
    bmp = gdi32.CreateCompatibleBitmap(screen_dc, BMP_W, BMP_H)
    old_bmp = gdi32.SelectObject(mem_dc, bmp)

    # Fill white background
    rect = wt.RECT(0, 0, BMP_W, BMP_H)
    white_brush = gdi32.GetStockObject(WHITE_BRUSH)
    user32.FillRect(mem_dc, ctypes.byref(rect), white_brush)

    # Create font
    family_w = ctypes.create_unicode_buffer("MS Gothic", 32)
    font = gdi32.CreateFontW(
        -PPEM,                    # nHeight (negative for em-based)
        0,                        # nWidth
        0, 0,                     # nEscapement, nOrientation
        400,                      # fnWeight (normal)
        0, 0, 0,                  # fdwItalic, fdwUnderline, fdwStrikeOut
        charset,                  # fdwCharSet
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY,
        DEFAULT_PITCH,
        family_w,
    )
    old_font = gdi32.SelectObject(mem_dc, font)

    # Set transparent background and black text
    gdi32.SetBkMode(mem_dc, TRANSPARENT)
    gdi32.SetTextColor(mem_dc, 0x000000)

    # Get TEXTMETRIC
    tm = TEXTMETRICW()
    gdi32.GetTextMetricsW(mem_dc, ctypes.byref(tm))

    # Get char advance via GetCharABCWidthsW (works for TrueType fonts)
    abc = ABC()
    code = ord(char)
    has_abc = bool(gdi32.GetCharABCWidthsW(mem_dc, code, code, ctypes.byref(abc)))

    # Get the actual font face name (resolved by GDI)
    face_buf = ctypes.create_unicode_buffer(64)
    face_len = gdi32.GetTextFaceW(mem_dc, 64, face_buf)
    actual_face = face_buf.value

    # Render char
    text_w = ctypes.create_unicode_buffer(char, 4)
    gdi32.TextOutW(mem_dc, x, y, text_w, len(char))

    # Read pixels via GetDIBits — get the bitmap data
    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wt.DWORD),
            ("biWidth", wt.LONG),
            ("biHeight", wt.LONG),
            ("biPlanes", wt.WORD),
            ("biBitCount", wt.WORD),
            ("biCompression", wt.DWORD),
            ("biSizeImage", wt.DWORD),
            ("biXPelsPerMeter", wt.LONG),
            ("biYPelsPerMeter", wt.LONG),
            ("biClrUsed", wt.DWORD),
            ("biClrImportant", wt.DWORD),
        ]

    class BITMAPINFO(ctypes.Structure):
        _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wt.DWORD * 3)]

    bi = BITMAPINFO()
    bi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bi.bmiHeader.biWidth = BMP_W
    bi.bmiHeader.biHeight = -BMP_H  # top-down
    bi.bmiHeader.biPlanes = 1
    bi.bmiHeader.biBitCount = 32
    bi.bmiHeader.biCompression = 0  # BI_RGB

    pixel_buf = (ctypes.c_ubyte * (BMP_W * BMP_H * 4))()
    gdi32.GetDIBits(mem_dc, bmp, 0, BMP_H, pixel_buf, ctypes.byref(bi), 0)

    # Find leftmost non-white pixel
    leftmost = None
    for px in range(BMP_W):
        for py in range(BMP_H):
            offset = (py * BMP_W + px) * 4
            b, g, r, a = pixel_buf[offset], pixel_buf[offset+1], pixel_buf[offset+2], pixel_buf[offset+3]
            # White is 255,255,255. Anything below 250 is "dark" (text).
            if r < 240 or g < 240 or b < 240:
                leftmost = px
                break
        if leftmost is not None:
            break

    # Cleanup
    gdi32.SelectObject(mem_dc, old_font)
    gdi32.DeleteObject(font)
    gdi32.SelectObject(mem_dc, old_bmp)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(0, screen_dc)

    return {
        "charset": charset,
        "tm_ascent": tm.tmAscent,
        "tm_descent": tm.tmDescent,
        "tm_height": tm.tmHeight,
        "tm_charset": tm.tmCharSet,
        "abc_a": abc.abcA if has_abc else None,
        "abc_b": abc.abcB if has_abc else None,
        "abc_c": abc.abcC if has_abc else None,
        "actual_face": actual_face,
        "render_x": x,
        "leftmost_pixel_x": leftmost,
        "lsb_px": leftmost - x if leftmost else None,
    }


def main():
    test_chars = ["□", "Ａ", "あ"]
    charsets = [
        ("DEFAULT_CHARSET", DEFAULT_CHARSET),
        ("SHIFTJIS_CHARSET", SHIFTJIS_CHARSET),
        ("ANSI_CHARSET", ANSI_CHARSET),
        ("SYMBOL_CHARSET", SYMBOL_CHARSET),
    ]

    print(f"Render with MS Gothic at PPEM={PPEM} (= 14pt @ 96dpi), x=20")
    print(f"{'Charset':<20} {'char':<4} {'face':<24} {'tmCharSet':>9} {'abcA':>5} {'abcB':>5} {'leftmost':>8} {'LSB_px':>6}")
    print("-" * 100)
    for ch in test_chars:
        for name, cs in charsets:
            r = render_char_with_charset(cs, ch)
            print(f"{name:<20} {ch:<4} {r['actual_face']:<24} {r['tm_charset']:>9} "
                  f"{r['abc_a'] if r['abc_a'] is not None else '-':>5} "
                  f"{r['abc_b'] if r['abc_b'] is not None else '-':>5} "
                  f"{r['leftmost_pixel_x'] if r['leftmost_pixel_x'] is not None else '-':>8} "
                  f"{r['lsb_px'] if r['lsb_px'] is not None else '-':>6}")
        print()


if __name__ == "__main__":
    main()
