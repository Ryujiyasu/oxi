"""Check if MS Gothic has actual glyph for U+25A1 (□) or falls back via GDI font linking.

GDI font linking (font fallback) on Japanese Windows automatically substitutes
glyphs from other fonts when the primary font lacks them. This may give MS Gothic
a different effective glyph for □ than its native one.

Use:
  - GetGlyphIndicesW with GGI_MARK_NONEXISTING_GLYPHS to see if MS Gothic has
    a native glyph for □
  - GetCharABCWidthsW to see GDI's reported metrics
  - GetCharWidthsI per-glyph for the linked font
"""
import ctypes
import ctypes.wintypes as wt
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

DEFAULT_CHARSET = 1
SHIFTJIS_CHARSET = 128
CLEARTYPE_QUALITY = 5
GGI_MARK_NONEXISTING_GLYPHS = 0x1


class ABC(ctypes.Structure):
    _fields_ = [("abcA", wt.LONG), ("abcB", wt.UINT), ("abcC", wt.LONG)]


def check_glyph(family: str, charset: int, char: str):
    screen_dc = user32.GetDC(0)
    mem_dc = gdi32.CreateCompatibleDC(screen_dc)
    family_w = ctypes.create_unicode_buffer(family, 64)
    PPEM = 19  # 14pt @ 96dpi
    font = gdi32.CreateFontW(
        -PPEM, 0, 0, 0, 400, 0, 0, 0,
        charset, 0, 0, CLEARTYPE_QUALITY, 0,
        family_w,
    )
    old = gdi32.SelectObject(mem_dc, font)

    # Resolved face name
    face_buf = ctypes.create_unicode_buffer(64)
    gdi32.GetTextFaceW(mem_dc, 64, face_buf)

    # Glyph index for char
    text_w = ctypes.create_unicode_buffer(char, 4)
    glyph_buf = (wt.WORD * 1)()
    res = gdi32.GetGlyphIndicesW(
        mem_dc, text_w, 1, glyph_buf, GGI_MARK_NONEXISTING_GLYPHS
    )
    glyph_idx = glyph_buf[0]
    has_glyph = glyph_idx != 0xFFFF

    # ABC widths
    abc = ABC()
    has_abc = bool(gdi32.GetCharABCWidthsW(
        mem_dc, ord(char), ord(char), ctypes.byref(abc)
    ))

    # ABC by glyph index — works for both native and linked
    abc_i = ABC()
    has_abc_i = bool(gdi32.GetCharABCWidthsI(
        mem_dc, glyph_idx, 1, None, ctypes.byref(abc_i)
    ))

    gdi32.SelectObject(mem_dc, old)
    gdi32.DeleteObject(font)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(0, screen_dc)

    return {
        "family_in": family,
        "charset_in": charset,
        "face_resolved": face_buf.value,
        "char": char,
        "codepoint": f"U+{ord(char):04X}",
        "glyph_idx": glyph_idx,
        "has_native_glyph": has_glyph,
        "abc_char": (abc.abcA, abc.abcB, abc.abcC) if has_abc else None,
        "abc_glyph": (abc_i.abcA, abc_i.abcB, abc_i.abcC) if has_abc_i else None,
    }


def main():
    fonts_to_test = [
        ("MS Gothic", DEFAULT_CHARSET),
        ("MS Gothic", SHIFTJIS_CHARSET),
        ("Yu Mincho", DEFAULT_CHARSET),
        ("Meiryo", DEFAULT_CHARSET),
        ("Arial", DEFAULT_CHARSET),
        ("Arial Unicode MS", DEFAULT_CHARSET),
    ]
    chars = ["□", "Ａ", "あ", "1"]

    print(f"{'family_in':<18} {'cs':>3} {'char':<4} {'face_resolved':<20} {'gidx':>5} {'native?':<7} {'abcA':>5} {'abcB':>5}")
    print("-" * 100)
    for family, cs in fonts_to_test:
        for ch in chars:
            r = check_glyph(family, cs, ch)
            native = "yes" if r["has_native_glyph"] else "NO"
            abcA = r["abc_char"][0] if r["abc_char"] else "-"
            abcB = r["abc_char"][1] if r["abc_char"] else "-"
            print(f"{family:<18} {cs:>3} {ch:<4} {r['face_resolved']:<20} {r['glyph_idx']:>5} {native:<7} {abcA:>5} {abcB:>5}")
        print()


if __name__ == "__main__":
    main()
