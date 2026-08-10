# -*- coding: utf-8 -*-
"""Probe: measure GDI float advances (GetCharABCWidthsFloatW) for Arial 18pt
and compare the per-line advance sum against Word's logical line width
(derived from spec5b Center/Right PDF render-truth).

Word logical line widths (pt):
  L1 'The quick brown fox jumps over'   ~254.04
  L5 'lazy dog.'                        ~72.0
"""
import ctypes
import ctypes.wintypes as wt

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32


class ABCFLOAT(ctypes.Structure):
    _fields_ = [("abcfA", ctypes.c_float),
                ("abcfB", ctypes.c_float),
                ("abcfC", ctypes.c_float)]


# 96 dpi screen DC (same as the renderer's dump path: GetDC(NULL)).
dc = user32.GetDC(None)
# -24px == 18pt at 96 dpi (matches create_font_for with scale=1.0 @96dpi).
font = gdi32.CreateFontW(-24, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, "Arial")
old = gdi32.SelectObject(dc, font)

def float_adv_sum_pt(s):
    """Sum of A+B+C float advances for each char in s, in POINTS (96dpi: unit = px, pt = px)."""
    total = 0.0
    for ch in s:
        cp = ord(ch)
        abc = ABCFLOAT()
        ok = gdi32.GetCharABCWidthsFloatW(dc, cp, cp, ctypes.byref(abc))
        if not ok:
            print("  FAILED char", repr(ch), cp)
            continue
        total += abc.abcfA + abc.abcfB + abc.abcfC
    # GetCharABCWidthsFloatW returns LOGICAL units; at 96 dpi 1 unit = 1/96 in = 0.75pt.
    return total * 0.75


lines = [
    "The quick brown fox jumps over",   # expect ~254.04
    "the lazy dog. The quick brown fox",  # expect ~267.75
    "jumps over the lazy dog. The",      # expect ~230.85
    "quick brown fox jumps over the",    # expect ~248.07
    "lazy dog.",                          # expect ~72.0
]

print("\n=== GDI float advance sums (pt) vs Word logical widths ===")
word_w = [254.04, 267.75, 230.85, 248.07, 72.0]
for s, ww in zip(lines, word_w):
    s_trim = s.rstrip()
    g = float_adv_sum_pt(s_trim)
    # integer GetTextExtentPoint32W too
    buf = (wt.WCHAR * (len(s_trim) + 1))(*s_trim)
    size = wt.SIZE()
    gdi32.GetTextExtentPoint32W(dc, buf, len(s_trim), ctypes.byref(size))
    g_int = size.cx * 0.75
    print(f"  '{s_trim}'  GDIfloat={g:7.2f}  GDIint={g_int:7.2f}  Word={ww:7.2f}  d_f={g-ww:+7.2f} d_i={g_int-ww:+7.2f}")

# Per-char GDI float advance for line 1 (compare against the PDF truth recorded
# in analyze_pptx_spec5b_rows.py for the Left block).
print("\n=== Per-char GDI float advances (pt) for L1 vs PDF (from rows dump) ===")
l1 = "The quick brown fox jumps over"
pdf_adv = {'T': 11.14, 'h': 10.01, 'e': 10.01, ' ': 4.91, 'q': 4.91, 'u': 10.01,
           'i': 9.92, 'c': 4.0, 'k': 9.0, 'b': 5.11, 'r': 10.01, 'o': 5.99,
           'w': 9.92, 'n': 12.6, 'f': 5.45, 'x': 10.01, 'j': 5.0, 'm': 10.01,
           'p': 14.99, 's': 9.92, 'v': 10.01}
for ch in l1:
    cp = ord(ch)
    abc = ABCFLOAT()
    gdi32.GetCharABCWidthsFloatW(dc, cp, cp, ctypes.byref(abc))
    g = (abc.abcfA + abc.abcfB + abc.abcfC) * 0.75
    p = pdf_adv.get(ch)
    d = "" if p is None else f"  d={g-p:+5.2f}"
    print(f"  '{ch}'  A={abc.abcfA*0.75:6.2f} B={abc.abcfB*0.75:6.2f} C={abc.abcfC*0.75:6.2f}  adv={g:6.2f}{d}")

gdi32.SelectObject(dc, old)
gdi32.DeleteObject(font)
user32.ReleaseDC(None, dc)
