"""
Ra: Calibri/Arial 小サイズ(7-12pt)の全文字幅テーブルをGDI計測
- Latin大文字・小文字・数字・記号の完全テーブル
- CJKフォールバック時(Calibri→MS UI Gothic)の文字幅
- 低SSIM文書で頻出する文字セット
"""
import ctypes, json, os

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


def gdi_widths(font_name, ppem, chars):
    hdc = user32.GetDC(0)
    hfont = gdi32.CreateFontW(-ppem, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, font_name)
    old = gdi32.SelectObject(hdc, hfont)
    widths = {}
    for ch in chars:
        sz = SIZE()
        gdi32.GetTextExtentPoint32W(hdc, ch, 1, ctypes.byref(sz))
        widths[ch] = sz.cx
    gdi32.SelectObject(hdc, old)
    gdi32.DeleteObject(hfont)
    user32.ReleaseDC(0, hdc)
    return widths


# Characters commonly found in low-SSIM documents
LATIN_UPPER = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
LATIN_LOWER = "abcdefghijklmnopqrstuvwxyz"
DIGITS = "0123456789"
SYMBOLS = " !\"#$%&'()*+,-./:;<=>?@[\\]^_{|}~"
JP_COMMON = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん"
JP_KATAKANA = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"
JP_PUNCT = "、。「」・ー（）〈〉《》【】〔〕"
KANJI_COMMON = "一二三四五六七八九十百千万円年月日時分第号条項"

ALL_CHARS = LATIN_UPPER + LATIN_LOWER + DIGITS + SYMBOLS + JP_COMMON[:10] + JP_KATAKANA[:5] + JP_PUNCT + KANJI_COMMON[:10]

results = []

# Font/size combos
configs = [
    ("Calibri", [7, 8, 9, 10, 10.5, 11, 12]),
    ("Arial", [7, 8, 9, 10, 11, 12]),
    ("Times New Roman", [9, 10, 10.5, 11, 12]),
    ("MS UI Gothic", [7, 8, 9, 10, 10.5, 11]),  # CJK fallback font
]

for font_name, sizes in configs:
    print(f"\n=== {font_name} ===")
    for fs in sizes:
        ppem = round(fs * 96.0 / 72.0)
        widths = gdi_widths(font_name, ppem, ALL_CHARS)

        entry = {"font": font_name, "size_pt": fs, "ppem": ppem, "widths": {}}
        for ch, w in widths.items():
            entry["widths"][f"U+{ord(ch):04X}"] = w

        results.append(entry)

        # Summary
        latin_w = [widths[c] for c in LATIN_UPPER if c in widths]
        digit_w = [widths[c] for c in DIGITS if c in widths]
        space_w = widths.get(' ', 0)

        print(f"  {fs}pt (ppem={ppem}): space={space_w}px, "
              f"A-Z avg={sum(latin_w)/len(latin_w):.1f}px, "
              f"0-9 avg={sum(digit_w)/len(digit_w):.1f}px")

# CJK fallback verification: Calibri CJK chars vs MS UI Gothic
print("\n\n=== CJK FALLBACK VERIFICATION ===")
print("Calibri CJK should fall back to MS UI Gothic\n")
cjk_test = JP_COMMON[:10] + KANJI_COMMON[:10]
for fs in [9, 10.5, 11]:
    ppem = round(fs * 96.0 / 72.0)
    cal_w = gdi_widths("Calibri", ppem, cjk_test)
    uig_w = gdi_widths("MS UI Gothic", ppem, cjk_test)

    matches = sum(1 for c in cjk_test if cal_w.get(c) == uig_w.get(c))
    total = len(cjk_test)
    mismatches = [(c, cal_w.get(c), uig_w.get(c)) for c in cjk_test if cal_w.get(c) != uig_w.get(c)]

    print(f"  {fs}pt: {matches}/{total} match", end="")
    if mismatches:
        print(f"  MISMATCHES: {mismatches[:5]}")
    else:
        print(" (ALL MATCH)")

# Save
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_calibri_arial_small.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
