"""
Ra: 小フォントサイズ(7-9pt)の文字幅精度をCOM計測で確定
- 低SSIM文書で多用される7pt, 8pt, 9ptの文字幅
- GDI pixel-round とOxi計算の差異
- 特にCalibri, MS Gothic, MS Mincho の小サイズ
"""
import win32com.client, json, os, ctypes
from ctypes import wintypes

# GDI functions for direct measurement
gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

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

class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


def gdi_measure_chars(font_name, font_size_pt, chars):
    """Measure character widths using GDI at 96 DPI."""
    hdc = user32.GetDC(0)
    ppem = round(font_size_pt * 96.0 / 72.0)

    hfont = gdi32.CreateFontW(
        -ppem, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, font_name
    )
    old_font = gdi32.SelectObject(hdc, hfont)

    results_list = []
    for ch in chars:
        sz = SIZE()
        gdi32.GetTextExtentPoint32W(hdc, ch, 1, ctypes.byref(sz))
        width_px = sz.cx
        width_pt = width_px * 72.0 / 96.0
        results_list.append({
            "char": ch,
            "code": hex(ord(ch)),
            "width_px": width_px,
            "width_pt": round(width_pt, 4),
        })

    gdi32.SelectObject(hdc, old_font)
    gdi32.DeleteObject(hfont)
    user32.ReleaseDC(0, hdc)

    return results_list


results = []

# Test characters
test_chars = list("ABCDabcd0123あいうアイウ一二三、。")

# Font/size combinations from low-SSIM documents
font_sizes = [
    ("Calibri", [7, 8, 9, 10, 10.5, 11]),
    ("MS Gothic", [7, 8, 9, 10, 10.5]),
    ("MS Mincho", [7, 8, 9, 10, 10.5]),
    ("Arial", [7, 8, 9, 10, 11]),
]

for font_name, sizes in font_sizes:
    print(f"\n=== {font_name} ===")
    for fs in sizes:
        ppem = round(fs * 96.0 / 72.0)
        widths = gdi_measure_chars(font_name, fs, test_chars)

        entry = {
            "font": font_name,
            "size_pt": fs,
            "ppem": ppem,
            "chars": widths,
        }
        results.append(entry)

        # Show summary
        latin_widths = [w for w in widths if ord(w["char"]) < 0x100]
        cjk_widths = [w for w in widths if ord(w["char"]) >= 0x3000]

        latin_avg = sum(w["width_px"] for w in latin_widths) / len(latin_widths) if latin_widths else 0
        cjk_avg = sum(w["width_px"] for w in cjk_widths) / len(cjk_widths) if cjk_widths else 0

        print(f"  {fs}pt (ppem={ppem}): Latin avg={latin_avg:.1f}px, CJK avg={cjk_avg:.1f}px")

        # Show individual chars for smallest size
        if fs == sizes[0]:
            for w in widths:
                print(f"    '{w['char']}' ({w['code']}): {w['width_px']}px = {w['width_pt']}pt")

# Save
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_small_font_charwidth.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")

# Analysis: check CJK fullwidth = ppem for MS Gothic
print("\n=== CJK FULLWIDTH = PPEM CHECK (MS Gothic) ===")
for entry in results:
    if entry["font"] == "MS Gothic":
        ppem = entry["ppem"]
        cjk_chars = [w for w in entry["chars"] if ord(w["char"]) >= 0x3040 and ord(w["char"]) < 0x9FFF]
        for w in cjk_chars[:3]:
            match = "OK" if w["width_px"] == ppem else f"DIFF(expected {ppem})"
            print(f"  {entry['size_pt']}pt: '{w['char']}' = {w['width_px']}px (ppem={ppem}) [{match}]")
