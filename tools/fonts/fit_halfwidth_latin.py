# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Fit the half-width Latin glyphs of OxiGothic / OxiMincho into their advance.

The two faces stand in for ＭＳ ゴシック / ＭＳ 明朝, whose ASCII glyphs advance
half an em. The shipped files kept Noto's PROPORTIONAL outlines under that
500-unit advance: OxiMincho «W» spans 12..1046, «M» 41..924, so a word set in
them overprints itself. This narrows every printable-ASCII glyph that advances
500 to at most 500 - 2 * MARGIN (x only, about its own centre) and centres it in
the advance, the way the ＭＳ faces draw their half-width Latin. Advances are
unchanged, so every metric table derived from the faces stays valid.

    python tools/fonts/fit_halfwidth_latin.py crates/oxidocs-fonts-jp/fonts/OxiMincho.ttf
"""
import sys
from fontTools.ttLib import TTFont

HALF = 500
MARGIN = 30


def fit(path):
    font = TTFont(path)
    upem = font['head'].unitsPerEm
    assert upem == 1000, upem
    glyf, hmtx = font['glyf'], font['hmtx']
    cmap = font.getBestCmap()
    done = 0
    seen = set()
    for cp in range(0x21, 0x7F):
        name = cmap.get(cp)
        if not name or name in seen:
            continue
        seen.add(name)
        adv, _ = hmtx[name]
        g = glyf[name]
        if adv != HALF or g.isComposite() or g.numberOfContours <= 0:
            continue
        g.recalcBounds(glyf)
        width = g.xMax - g.xMin
        target = HALF - 2 * MARGIN
        scale = min(1.0, target / width) if width > 0 else 1.0
        centre = (g.xMin + g.xMax) / 2.0
        coords = g.coordinates
        for i in range(len(coords)):
            x, y = coords[i]
            coords[i] = (round(HALF / 2.0 + (x - centre) * scale), y)
        if hasattr(g, 'program'):
            g.program.fromBytecode(b'')
        g.recalcBounds(glyf)
        hmtx[name] = (adv, g.xMin)
        done += 1
    font.save(path)
    return done


if __name__ == '__main__':
    for p in sys.argv[1:]:
        print(p, fit(p), 'glyphs fitted')
