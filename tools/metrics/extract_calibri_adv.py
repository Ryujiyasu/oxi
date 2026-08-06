# Extract Calibri ASCII 32..=126 hmtx design advances for font_adv.rs.
# Mirrors gen_pptx_font_adv.py's format (95 entries, code-point order).
import sys
sys.stdout.reconfigure(encoding="utf-8")
from fontTools.ttLib import TTFont

p = r"C:\Windows\Fonts\calibri.ttf"
f = TTFont(p)
upm = f["head"].unitsPerEm
hmtx = f["hmtx"].metrics
cmap = f.getBestCmap()

print("upm", upm)
vals = []
for code in range(32, 127):
    g = cmap.get(code)
    if g is None or g not in hmtx:
        print("MISSING", code, chr(code), g)
        vals.append(0.0)
    else:
        vals.append(round(hmtx[g][0] / upm, 5))

for i in range(0, len(vals), 10):
    print("    " + ", ".join(f"{v:.5f}" for v in vals[i : i + 10]) + ",")
print("count", len(vals))
