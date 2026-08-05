# Generate an hmtx (design advance) width table for the fonts used by the
# PowerPoint (oxislides) renderer, in EM units (advance / unitsPerEm).
#
# Why: PowerPoint's PDF export places glyphs at their DESIGN advance
# (the TrueType hmtx value), NOT at GDI's hinted / integer-pixel snapped
# advance. A line's logical width in Word is the sum of the hmtx advances of
# its visible characters (trailing spaces excluded). GDI
# (GetCharABCWidthsFloatW / GetTextExtentPoint32W) returns hinted,
# pixel-rounded values (multiples of 1px @96dpi = 0.75pt), so a line that
# Word measures at 254.04pt comes out 255.75pt in GDI (+1.71).
#
# The generated table is embedded in tools/oxi-pptx-renderer/src/font_adv.rs
# and consulted for line-width / character-position computation.
#
# Usage:
#   python tools/metrics/gen_pptx_font_adv.py > /tmp/font_adv_table.json
import json
from fontTools.ttLib import TTFont

FONTS = {
    "arial": r"C:\Windows\Fonts\arial.ttf",
    "arialbd": r"C:\Windows\Fonts\arialbd.ttf",
}


def build(font_path):
    f = TTFont(font_path)
    upm = f["head"].unitsPerEm
    hmtx = f["hmtx"]
    cmap = f.getBestCmap()
    table = {}
    # ASCII 32..126 (printable) covers the repro text + common punctuation.
    for cp in range(32, 127):
        gname = cmap.get(cp)
        if gname is None:
            continue
        aw = hmtx[gname][0] if gname in hmtx.metrics else 0
        table[chr(cp)] = aw / upm
    return table


out = {name: build(path) for name, path in FONTS.items()}
print(json.dumps(out, ensure_ascii=False, sort_keys=True, indent=1))
