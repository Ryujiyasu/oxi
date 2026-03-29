"""
Ra: CJK偶数px丸めの範囲を確定
- MS Gothic/Mincho以外のCJKフォントでも偶数丸めか？
- 半角文字は fullwidth/2 か？ それとも独自丸め？
- Yu Gothic, Yu Mincho, Meiryo の挙動
- ppem偶数の場合と奇数の場合の全パターン
"""
import ctypes
from ctypes import wintypes
import json, os

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


def gdi_width(font_name, font_size_pt, char):
    hdc = user32.GetDC(0)
    ppem = round(font_size_pt * 96.0 / 72.0)
    hfont = gdi32.CreateFontW(-ppem, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, font_name)
    old = gdi32.SelectObject(hdc, hfont)
    sz = SIZE()
    gdi32.GetTextExtentPoint32W(hdc, char, 1, ctypes.byref(sz))
    gdi32.SelectObject(hdc, old)
    gdi32.DeleteObject(hfont)
    user32.ReleaseDC(0, hdc)
    return sz.cx


results = []

# Test all CJK fonts at ppem 7-20
fonts = ["MS Gothic", "MS Mincho", "MS PGothic", "MS PMincho",
         "Yu Gothic", "Yu Mincho", "Meiryo", "HGGothicM"]

test_char_fw = "\u3042"  # あ (fullwidth hiragana)
test_char_hw = "A"       # halfwidth
test_char_kanji = "\u4e00"  # 一

print("=== CJK FULLWIDTH ROUNDING PATTERN ===\n")
print(f"{'Font':<15} {'ppem':>4} {'full_px':>7} {'ceil_even':>9} {'match':>5} {'half_px':>7} {'fw/2':>4}")
print("-" * 70)

for font in fonts:
    font_data = {"font": font, "measurements": []}
    for ppem in range(7, 21):
        fs_pt = ppem * 72.0 / 96.0
        fw = gdi_width(font, fs_pt, test_char_fw)
        hw = gdi_width(font, fs_pt, test_char_hw)
        kanji = gdi_width(font, fs_pt, test_char_kanji)

        ceil_even = (ppem + 1) & ~1  # ceil to even
        match = fw == ceil_even

        font_data["measurements"].append({
            "ppem": ppem, "fullwidth_px": fw, "halfwidth_px": hw,
            "kanji_px": kanji, "ceil_even": ceil_even, "match": match,
        })

        if not match or ppem <= 14:
            print(f"{font:<15} {ppem:>4} {fw:>7} {ceil_even:>9} {'OK' if match else 'NG':>5} {hw:>7} {fw//2:>4}")

    results.append(font_data)

# Detailed analysis for proportional CJK fonts
print("\n\n=== PROPORTIONAL CJK (MS PGothic, Yu Gothic, Meiryo) ===\n")
prop_fonts = ["MS PGothic", "Yu Gothic", "Meiryo"]
test_chars = list("あいうえおアイウ一二三ABab12、。（）")

for font in prop_fonts:
    print(f"\n--- {font} ---")
    for fs in [9, 10.5, 11]:
        ppem = round(fs * 96.0 / 72.0)
        print(f"  {fs}pt (ppem={ppem}):")
        for ch in test_chars:
            w = gdi_width(font, fs, ch)
            code = hex(ord(ch))
            is_fw = ord(ch) >= 0x3000
            marker = "*" if is_fw and w != ppem else ""
            print(f"    '{ch}'({code}): {w}px{marker}", end="")
        print()

# Check the specific pattern: is it ceil_even or something else?
print("\n\n=== FORMULA VERIFICATION ===\n")
for font in ["MS Gothic", "MS Mincho"]:
    print(f"{font}:")
    all_match_ceil_even = True
    all_match_ppem = True
    for ppem in range(5, 30):
        fs_pt = ppem * 72.0 / 96.0
        fw = gdi_width(font, fs_pt, test_char_fw)
        ceil_even = (ppem + 1) & ~1
        if fw != ceil_even:
            all_match_ceil_even = False
            print(f"  ppem={ppem}: fw={fw}, ceil_even={ceil_even}, ppem={ppem}")
        if fw != ppem:
            all_match_ppem = False
    print(f"  ceil_even formula: {'ALL MATCH' if all_match_ceil_even else 'has mismatches'}")
    print(f"  ppem direct: {'ALL MATCH' if all_match_ppem else 'has mismatches'}")

# Save
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_cjk_evenround.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
