"""
Ra: GDI幅テーブル生成 — 全主要フォント × 全ppem × 全Latin+CJK文字
Oxiに組み込むための事前計算テーブル
"""
import ctypes, json, os

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


def gdi_widths_batch(font_name, ppem, codepoints):
    hdc = user32.GetDC(0)
    hfont = gdi32.CreateFontW(-ppem, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, font_name)
    old = gdi32.SelectObject(hdc, hfont)
    widths = {}
    for cp in codepoints:
        ch = chr(cp)
        sz = SIZE()
        gdi32.GetTextExtentPoint32W(hdc, ch, 1, ctypes.byref(sz))
        widths[cp] = sz.cx
    gdi32.SelectObject(hdc, old)
    gdi32.DeleteObject(hfont)
    user32.ReleaseDC(0, hdc)
    return widths


def gdi_widths_bold(font_name, ppem, codepoints):
    hdc = user32.GetDC(0)
    hfont = gdi32.CreateFontW(-ppem, 0, 0, 0, 700, 0, 0, 0, 0, 0, 0, 0, 0, font_name)
    old = gdi32.SelectObject(hdc, hfont)
    widths = {}
    for cp in codepoints:
        ch = chr(cp)
        sz = SIZE()
        gdi32.GetTextExtentPoint32W(hdc, ch, 1, ctypes.byref(sz))
        widths[cp] = sz.cx
    gdi32.SelectObject(hdc, old)
    gdi32.DeleteObject(hfont)
    user32.ReleaseDC(0, hdc)
    return widths


# Character ranges to measure
# Basic Latin (U+0020-U+007E)
BASIC_LATIN = list(range(0x0020, 0x007F))
# Latin-1 Supplement (U+00A0-U+00FF) — common accented chars
LATIN1_SUP = list(range(0x00A0, 0x0100))
# General punctuation subset
GEN_PUNCT = list(range(0x2000, 0x2070))
# CJK punctuation (U+3000-U+303F)
CJK_PUNCT = list(range(0x3000, 0x3040))
# Hiragana (U+3040-U+309F)
HIRAGANA = list(range(0x3040, 0x30A0))
# Katakana (U+30A0-U+30FF)
KATAKANA = list(range(0x30A0, 0x3100))
# Fullwidth Latin (U+FF00-U+FF5E)
FULLWIDTH = list(range(0xFF00, 0xFF5F))
# Common Kanji subset (most frequent ~200)
KANJI_COMMON = list(range(0x4E00, 0x4E60)) + list(range(0x5000, 0x5030)) + \
               list(range(0x6700, 0x6730)) + list(range(0x7530, 0x7560))

ALL_CODEPOINTS = sorted(set(
    BASIC_LATIN + LATIN1_SUP + GEN_PUNCT + CJK_PUNCT +
    HIRAGANA + KATAKANA + FULLWIDTH + KANJI_COMMON
))

print(f"Total codepoints to measure: {len(ALL_CODEPOINTS)}")

# Fonts and ppem ranges
# Focus on ppem values that have mismatches (odd ppem for UPM=2048 fonts)
FONTS = [
    ("Calibri", False),
    ("Calibri", True),  # Bold
    ("Arial", False),
    ("Arial", True),
    ("Times New Roman", False),
    ("Times New Roman", True),
    ("Century", False),
    ("Cambria", False),
]

# ppem 7-20 covers 5.25pt to 15pt at 96 DPI
PPEMS = list(range(7, 21))

results = {}

for font_name, bold in FONTS:
    label = f"{font_name}{' Bold' if bold else ''}"
    print(f"\n=== {label} ===")
    font_data = {}

    for ppem in PPEMS:
        if bold:
            widths = gdi_widths_bold(font_name, ppem, ALL_CODEPOINTS)
        else:
            widths = gdi_widths_batch(font_name, ppem, ALL_CODEPOINTS)

        # Only store non-zero widths
        non_zero = {str(cp): w for cp, w in widths.items() if w > 0}
        font_data[str(ppem)] = non_zero

        # Count how many differ from round(advance * ppem / upm)
        fs_pt = ppem * 72.0 / 96.0
        print(f"  ppem={ppem} ({fs_pt:.1f}pt): {len(non_zero)} chars measured")

    results[label] = font_data

# Also do MS UI Gothic (CJK fallback)
for font_name in ["MS UI Gothic"]:
    print(f"\n=== {font_name} ===")
    font_data = {}
    for ppem in PPEMS:
        widths = gdi_widths_batch(font_name, ppem, ALL_CODEPOINTS)
        non_zero = {str(cp): w for cp, w in widths.items() if w > 0}
        font_data[str(ppem)] = non_zero
        fs_pt = ppem * 72.0 / 96.0
        print(f"  ppem={ppem} ({fs_pt:.1f}pt): {len(non_zero)} chars measured")
    results[font_name] = font_data

# Save
out_dir = os.path.join(os.path.dirname(__file__), '..', '..',
    'crates', 'oxidocs-core', 'src', 'font', 'data')

out_path = os.path.join(out_dir, 'gdi_width_overrides.json')
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, separators=(',', ':'))

file_size = os.path.getsize(out_path) / 1024
print(f"\nSaved to {out_path} ({file_size:.1f} KB)")

# Also save a compact version with only differences from round(advance*ppem/upm)
# Load Oxi metrics for comparison
oxi_path = os.path.join(out_dir, 'font_metrics_compact.json')
with open(oxi_path) as f:
    oxi_metrics = json.load(f)

oxi_map = {}
for fm in oxi_metrics:
    family = fm.get("family", "")
    oxi_map[family] = {"upm": fm["units_per_em"], "widths": fm.get("widths", {})}

diff_results = {}
total_diffs = 0
for label, ppem_data in results.items():
    base_name = label.replace(" Bold", "")
    oxi_fm = oxi_map.get(base_name)
    if not oxi_fm:
        continue

    upm = oxi_fm["upm"]
    oxi_widths = oxi_fm["widths"]
    font_diffs = {}

    for ppem_str, gdi_widths_map in ppem_data.items():
        ppem = int(ppem_str)
        ppem_diffs = {}

        for cp_str, gdi_w in gdi_widths_map.items():
            oxi_advance = oxi_widths.get(cp_str)
            if oxi_advance is not None:
                oxi_px = round(oxi_advance * ppem / upm)
                if oxi_px != gdi_w:
                    ppem_diffs[cp_str] = gdi_w
                    total_diffs += 1

        if ppem_diffs:
            font_diffs[ppem_str] = ppem_diffs

    if font_diffs:
        diff_results[label] = font_diffs

diff_path = os.path.join(out_dir, 'gdi_pixel_overrides.json')
with open(diff_path, 'w', encoding='utf-8') as f:
    json.dump(diff_results, f, separators=(',', ':'))

diff_size = os.path.getsize(diff_path) / 1024
print(f"Diff-only overrides: {diff_path} ({diff_size:.1f} KB, {total_diffs} overrides)")
