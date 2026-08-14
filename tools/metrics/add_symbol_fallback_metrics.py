# -*- coding: utf-8 -*-
"""Add the symbol FALLBACK faces + a glyph-coverage bitmap to the metrics table.

S1115 shipped the rule that a Latin document's ambiguous-width symbols resolve
through the ASCII font.  Its residual 1 is that Word does NOT stop there: when
the ascii font has no glyph for the codepoint, Word falls back to another face
and the line takes THAT face's height.  `_pb_symline_gen.py` measures it (18pt
symbol beside 10pt text, 8 copies per arm):

    Arial   (1) / (diamond)     21.094   = Cambria Math    1.1719 x 18
    Arial   ballot / check / star 24.000  = Segoe UI Symbol 1.3333 x 18
    Calibri black square         20.438   = Courier New     1.1354 x 18
    Cambria ballot / check / star 23.906  = Segoe UI Symbol (0.094 = device)

Cross-referencing those against the faces' cmaps gives ONE ordered chain that
explains all 39 arms:

    the run font itself -> Courier New -> Cambria Math -> Segoe UI Symbol

The two cases that pin the ORDER are Calibri black-square (Courier New has it,
so Courier New wins over Cambria Math which also has it) and Calibri diamond
(Courier New lacks it, so Cambria Math wins).  Everything else is consistent
but not discriminating -- if a later measurement contradicts the order, it is
those two arms that have to be re-measured first.

★★CAMBRIA MATH IS A TRAP. Its OS/2 win metrics are inflated for stretchy
math glyphs -- win sum = 5.5801 em against an hhea sum of 1.1733 -- so the
registry's usual max(hhea, win) natural height would give a 100pt line at
18pt. Word's measured fallback line through that face is 21.094 / 18 =
1.1719 em, i.e. it follows **hhea**, not max(hhea, win). The other two
faces have hhea == win and cannot tell the two models apart, so Cambria
Math is the only arm that pins this: the fallback line height is the
fallback face's HHEA sum (then device-rounded, which accounts for the
residual +0.03..+0.06pt against Word).

This script is data-only, in the shape of add_garamond_metrics.py and its
siblings: it adds the two missing faces, and gives EVERY entry a `sym_coverage`
hex bitmap over S1115's ambiguous ranges so the layout side can ask "does this
face have this codepoint" without shipping a font file or a full cmap.

usage: python tools/metrics/add_symbol_fallback_metrics.py [--dry-run]
"""

import json
import os
import sys

JSON = os.path.join(
    "crates", "oxidocs-core", "src", "font", "data", "font_metrics_compact.json"
)

# The faces the fallback chain needs.  Courier New is already in the table.
FACES = [
    ("Segoe UI Symbol", r"C:\Windows\Fonts\seguisym.ttf", 0),
    ("Cambria Math", r"C:\Windows\Fonts\cambria.ttc", 1),
]
# Where to look for each registry family's file, so coverage can be computed.
FILE_FOR = {
    "Arial": ("arial.ttf", 0), "Arial Bold": ("arialbd.ttf", 0),
    "Arial Narrow": ("ARIALN.TTF", 0), "Calibri": ("calibri.ttf", 0),
    "Calibri Light": ("calibril.ttf", 0), "Cambria": ("cambria.ttc", 0),
    "Cambria Math": ("cambria.ttc", 1), "Courier New": ("cour.ttf", 0),
    "Times New Roman": ("times.ttf", 0), "Segoe UI": ("segoeui.ttf", 0),
    "Segoe UI Symbol": ("seguisym.ttf", 0), "Verdana": ("verdana.ttf", 0),
    "Tahoma": ("tahoma.ttf", 0), "Georgia": ("georgia.ttf", 0),
    "Trebuchet MS": ("trebuc.ttf", 0), "Garamond": ("GARA.TTF", 0),
    "Book Antiqua": ("BKANT.TTF", 0), "Century Gothic": ("GOTHIC.TTF", 0),
    "Bookman Old Style": ("BOOKOS.TTF", 0), "Aptos": ("Aptos.ttf", 0),
    "Franklin Gothic Book": ("FRABK.TTF", 0), "Comic Sans MS": ("comic.ttf", 0),
    "Palatino Linotype": ("pala.ttf", 0), "Bell MT": ("BELL.TTF", 0),
    "Open Sans": ("OpenSans-Regular.ttf", 0), "Symbol": ("symbol.ttf", 0),
    "Wingdings": ("wingding.ttf", 0), "Webdings": ("webdings.ttf", 0),
    "MS Gothic": ("msgothic.ttc", 0), "MS Mincho": ("msmincho.ttc", 0),
    "Meiryo": ("meiryo.ttc", 0), "Yu Gothic": ("YuGothR.ttc", 0),
}
# S1115's ambiguous ranges, verbatim -- the codepoints the shipped rule routes
# to the ascii font, hence exactly the set where this fallback can bite.
RANGES = [(0x2010, 0x2044), (0x2190, 0x22FF), (0x2460, 0x24FF), (0x2500, 0x27BF)]
FONTDIR = r"C:\Windows\Fonts"


def all_codepoints():
    out = []
    for a, b in RANGES:
        out.extend(range(a, b + 1))
    return out


def coverage_hex(path, index):
    """One bit per codepoint of RANGES, LSB-first within each byte."""
    from fontTools.ttLib import TTFont

    f = TTFont(path, fontNumber=index, lazy=True)
    cmap = set(f.getBestCmap())
    f.close()
    cps = all_codepoints()
    bits = bytearray((len(cps) + 7) // 8)
    for i, cp in enumerate(cps):
        if cp in cmap:
            bits[i >> 3] |= 1 << (i & 7)
    return bits.hex()


def extract(path, index, codepoints):
    from fontTools.ttLib import TTFont

    f = TTFont(path, fontNumber=index, lazy=True)
    head, hhea, os2 = f["head"], f["hhea"], f["OS/2"]
    upm = head.unitsPerEm
    cmap = f.getBestCmap()
    hmtx = f["hmtx"]
    widths = {}
    for cp in codepoints:
        g = cmap.get(int(cp))
        if g is not None:
            widths[str(cp)] = int(round(hmtx[g][0] * 2048 / upm))
    e = {
        "family": None,
        "units_per_em": 2048,
        "ascender": int(round(hhea.ascent * 2048 / upm)),
        "descender": int(round(hhea.descent * 2048 / upm)),
        "line_gap": int(round(hhea.lineGap * 2048 / upm)),
        "win_ascent": int(round(os2.usWinAscent * 2048 / upm)),
        "win_descent": int(round(os2.usWinDescent * 2048 / upm)),
        "typo_ascender": int(round(os2.sTypoAscender * 2048 / upm)),
        "typo_descender": int(round(os2.sTypoDescender * 2048 / upm)),
        "typo_line_gap": int(round(os2.sTypoLineGap * 2048 / upm)),
        "widths": widths,
    }
    f.close()
    return e


def main():
    dry = "--dry-run" in sys.argv
    data = json.load(open(JSON, encoding="utf-8"))
    ref = [x for x in data if x.get("family") == "Verdana"][0]
    codepoints = sorted(int(k) for k in ref["widths"])
    have = {x.get("family") for x in data}

    for fam, path, idx in FACES:
        if fam in have:
            print("already present: %s" % fam)
            continue
        if not os.path.exists(path):
            print("MISSING FONT FILE: %s" % path)
            return 1
        e = extract(path, idx, codepoints)
        e["family"] = fam
        nat = max(e["ascender"] + abs(e["descender"]) + e["line_gap"],
                  e["win_ascent"] + e["win_descent"]) / 2048.0
        print("+ %-18s natural=%.6f em  widths=%d" % (fam, nat, len(e["widths"])))
        data.append(e)

    n_cov = 0
    for e in data:
        fam = e.get("family")
        spec = FILE_FOR.get(fam)
        if not spec:
            continue
        fn, idx = spec
        p = os.path.join(FONTDIR, fn)
        if not os.path.exists(p):
            continue
        try:
            e["sym_coverage"] = coverage_hex(p, idx)
            n_cov += 1
        except Exception as exc:  # noqa: BLE001
            print("  coverage failed for %s: %s" % (fam, exc))
    print("sym_coverage written for %d / %d entries (%d codepoints each)"
          % (n_cov, len(data), len(all_codepoints())))

    if dry:
        print("(dry run, not written)")
        return 0
    with open(JSON, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, separators=(",", ":"))
    print("wrote %s (%d entries)" % (JSON, len(data)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
