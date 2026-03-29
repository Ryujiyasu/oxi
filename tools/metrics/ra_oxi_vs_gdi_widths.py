"""
Ra: Oxi font_metrics_compact.json の文字幅 vs GDI実測値の差異を特定
- 低SSIM文書で使われるフォント/サイズの文字幅比較
- round(advance * ppem / upm) と GDI GetTextExtentPoint32W の差
"""
import ctypes, json, os

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


def gdi_width(font_name, ppem, char):
    hdc = user32.GetDC(0)
    hfont = gdi32.CreateFontW(-ppem, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, font_name)
    old = gdi32.SelectObject(hdc, hfont)
    sz = SIZE()
    gdi32.GetTextExtentPoint32W(hdc, char, 1, ctypes.byref(sz))
    gdi32.SelectObject(hdc, old)
    gdi32.DeleteObject(hfont)
    user32.ReleaseDC(0, hdc)
    return sz.cx


# Load Oxi font metrics
metrics_path = os.path.join(os.path.dirname(__file__), '..', '..',
    'crates', 'oxidocs-core', 'src', 'font', 'data', 'font_metrics_compact.json')
with open(metrics_path, 'r') as f:
    oxi_metrics = json.load(f)

results = []

# Test fonts that appear in low-SSIM documents
test_fonts = {
    "Calibri": "Calibri",
    "MS Gothic": "MS Gothic",
    "MS Mincho": "MS Mincho",
    "Arial": "Arial",
}

test_chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 "
test_sizes = [9, 10.5, 11]

for oxi_name, gdi_name in test_fonts.items():
    # Find in Oxi metrics
    font_data = None
    for fm in oxi_metrics:
        if fm.get("family") == oxi_name:
            font_data = fm
            break

    if not font_data:
        print(f"  {oxi_name}: NOT FOUND in Oxi metrics")
        continue

    upm = font_data.get("units_per_em", 2048)
    widths_map = font_data.get("widths", {})

    print(f"\n=== {oxi_name} (UPM={upm}) ===")

    for fs in test_sizes:
        ppem = round(fs * 96.0 / 72.0)

        mismatches = []
        total = 0
        max_diff = 0

        for ch in test_chars:
            cp = str(ord(ch))
            oxi_advance = widths_map.get(cp)
            if oxi_advance is None:
                continue

            total += 1
            oxi_px = round(oxi_advance * ppem / upm)
            gdi_px = gdi_width(gdi_name, ppem, ch)

            diff = oxi_px - gdi_px
            if diff != 0:
                mismatches.append({
                    "char": ch,
                    "code": f"U+{ord(ch):04X}",
                    "oxi_advance": oxi_advance,
                    "oxi_px": oxi_px,
                    "gdi_px": gdi_px,
                    "diff": diff,
                })
                max_diff = max(max_diff, abs(diff))

        results.append({
            "font": oxi_name,
            "size_pt": fs,
            "ppem": ppem,
            "total_chars": total,
            "mismatch_count": len(mismatches),
            "max_diff_px": max_diff,
            "mismatches": mismatches,
        })

        print(f"  {fs}pt (ppem={ppem}): {len(mismatches)}/{total} mismatches, max_diff={max_diff}px")
        if mismatches:
            for m in mismatches[:10]:
                print(f"    '{m['char']}'({m['code']}): oxi={m['oxi_px']}px, gdi={m['gdi_px']}px (diff={m['diff']})")

# Save
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_oxi_vs_gdi.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")

# Summary
print("\n=== OVERALL MISMATCH SUMMARY ===")
for r in results:
    if r["mismatch_count"] > 0:
        print(f"  {r['font']} {r['size_pt']}pt: {r['mismatch_count']} mismatches")
