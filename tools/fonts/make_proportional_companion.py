# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Build OxiGothicP / OxiMinchoP: the proportional companions of OxiGothic / OxiMincho.

OxiGothic and OxiMincho set printable ASCII at half an em, as ＭＳ ゴシック /
ＭＳ 明朝 draw it (see fit_halfwidth_latin.py), and kana, CJK punctuation and
the full-width forms at a full em. A family whose Latin is proportional --
游ゴシック, メイリオ, ＭＳ Ｐゴシック, 游明朝… -- drawn with them comes out in a
typewriter face with a gap after every word, and one whose kana are
proportional too -- ＭＳ Ｐゴシック, ＭＳ Ｐ明朝 -- has each 「（」「・」 or kana
drawn a full em wide in a cell the layout measured narrower, running into the
next character.

These companion faces carry
  * printable ASCII, with Noto's own outlines at Noto's own advances, and
  * U+3000-30FF (CJK punctuation, kana) and U+FF01-FF60 (full-width forms)
    with Noto's 'palt' (proportional alternate widths) applied: each glyph moved
    by its XPlacement and its advance cut by its XAdvance, so （ 」 ・ come out
    half an em wide with their ink where a proportional face puts it,
plus the vertical metrics of the face they accompany, so a renderer can take
those characters of a proportional family from them and everything else from
the main face. Which characters a family takes from here is the renderer's to
say, from the widths measured for that family.

The source is the variable Noto Sans JP 2.004 / Noto Serif JP 2.003 from Google
Fonts, instanced at Regular (400). Pass the pre-fit main face (the file before
fit_halfwidth_latin.py ran, from git) as a fifth argument to check that every
ASCII outline is the one the main face was cut with:

    python tools/fonts/make_proportional_companion.py 'NotoSansJP[wght].ttf' \\
        crates/oxidocs-fonts-jp/fonts/OxiGothic.ttf crates/oxidocs-fonts-jp/fonts/OxiGothicP.ttf OxiGothicP
    python tools/fonts/make_proportional_companion.py 'NotoSerifJP[wght].ttf' \\
        crates/oxidocs-fonts-jp/fonts/OxiMincho.ttf crates/oxidocs-fonts-jp/fonts/OxiMinchoP.ttf OxiMinchoP
"""
import sys
from fontTools.ttLib import TTFont
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.pens.transformPen import TransformPen
from fontTools.varLib import instancer
from fontTools import subset

ASCII = range(0x20, 0x7F)
PROPORTIONAL = [*range(0x3000, 0x3100), *range(0xFF01, 0xFF61)]


def palt(font):
    """Glyph name -> (XPlacement, XAdvance) of the font's 'palt' feature."""
    gpos = font['GPOS'].table
    lookups = sorted({i for fr in gpos.FeatureList.FeatureRecord if fr.FeatureTag == 'palt'
                      for i in fr.Feature.LookupListIndex})
    out = {}
    for i in lookups:
        lookup = gpos.LookupList.Lookup[i]
        for st in lookup.SubTable:
            if lookup.LookupType == 9:
                st = st.ExtSubTable
            assert st.LookupType == 1, f'palt lookup {i} is not single positioning'
            glyphs = st.Coverage.glyphs
            values = [st.Value] * len(glyphs) if st.Format == 1 else st.Value
            for g, v in zip(glyphs, values):
                out.setdefault(g, (getattr(v, 'XPlacement', 0) or 0, getattr(v, 'XAdvance', 0) or 0))
    return out


def build(noto_path, main_path, out_path, family):
    font = instancer.instantiateVariableFont(TTFont(noto_path), {'wght': 400})
    adjust = palt(font)
    cmap = font.getBestCmap()
    wanted = [cp for cp in (*ASCII, *PROPORTIONAL) if cp in cmap]

    options = subset.Options()
    options.layout_features = ['kern']
    options.name_IDs = ['*']
    options.name_languages = ['*']
    options.notdef_outline = True
    options.hinting = False
    sub = subset.Subsetter(options)
    sub.populate(unicodes=wanted)
    sub.subset(font)

    # Bake 'palt' into the kana and punctuation: the glyph moves by its
    # placement and its advance shrinks. Each is redrawn as a simple glyph, so
    # a component it shares with another is never moved twice.
    glyf, hmtx = font['glyf'], font['hmtx']
    glyph_set = font.getGlyphSet()
    cmap = font.getBestCmap()
    targets = sorted({cmap[cp] for cp in PROPORTIONAL if cp in cmap and cmap[cp] in adjust})
    moved = {}
    for name in targets:
        dx, _ = adjust[name]
        pen = TTGlyphPen(glyph_set)
        glyph_set[name].draw(TransformPen(pen, (1, 0, 0, 1, dx, 0)))
        moved[name] = pen.glyph()
    for name, glyph in moved.items():
        glyf[name] = glyph
        glyph.recalcBounds(glyf)
        advance, _ = hmtx[name]
        hmtx[name] = (advance + adjust[name][1], glyph.xMin if glyph.numberOfContours else 0)

    main = TTFont(main_path)
    # The main face's line metrics, so a line mixing the two sits on one baseline.
    for attr in ('ascent', 'descent', 'lineGap'):
        setattr(font['hhea'], attr, getattr(main['hhea'], attr))
    for attr in ('sTypoAscender', 'sTypoDescender', 'sTypoLineGap', 'usWinAscent', 'usWinDescent', 'fsSelection'):
        setattr(font['OS/2'], attr, getattr(main['OS/2'], attr))
    font['OS/2'].usWeightClass = 400

    names = font['name']
    for rec in list(names.names):
        if rec.nameID in (1, 3, 4, 6, 16, 17, 21, 22, 25):
            names.removeNames(nameID=rec.nameID)
    for pid, eid, lid in ((3, 1, 0x409), (1, 0, 0)):
        names.setName(family, 1, pid, eid, lid)
        names.setName('Regular', 2, pid, eid, lid)
        names.setName(f'{family}-Regular', 3, pid, eid, lid)
        names.setName(family, 4, pid, eid, lid)
        names.setName(family, 6, pid, eid, lid)
    for tag in ('STAT', 'fvar', 'gvar', 'HVAR', 'MVAR', 'avar', 'VORG', 'vhea', 'vmtx'):
        if tag in font:
            del font[tag]
    font.save(out_path)
    return len(moved)


def check(out_path, prefit_main):
    """Every ASCII outline matches the main face's own before it was fitted."""
    new, old = TTFont(out_path), TTFont(prefit_main)
    ncm, ocm = new.getBestCmap(), old.getBestCmap()
    for cp in ASCII:
        a, b = new['glyf'][ncm[cp]], old['glyf'][ocm[cp]]
        if a.numberOfContours > 0:
            assert list(a.coordinates) == list(b.coordinates), chr(cp)


if __name__ == '__main__':
    noto, main, out, family = sys.argv[1:5]
    count = build(noto, main, out, family)
    if len(sys.argv) > 5:
        check(out, sys.argv[5])
    print(out, 'written,', count, "glyphs set at 'palt' widths")
