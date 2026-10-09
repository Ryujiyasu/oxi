# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Export numeric glyph geometry from the measured face, without outlines."""

import struct

from fontTools.pens.boundsPen import BoundsPen


def face_geometry(font, face_key, *, ssty_source=None, stored_glyf=True):
    """Retain glyph identity, bounds, ssty, and MATH positioning metadata.

    Original TrueType faces use the glyf header's design bounds, including
    composite glyphs. Instantiated fonts must read their changed glyf table;
    their reader still contains the original program. CFF bounds are measured
    through a pen in memory. Neither charstrings nor outlines enter the output.
    Unsupported ssty lookups fail explicitly rather than losing script data.
    """
    if not face_key:
        raise ValueError('Glyph geometry requires a measured face identity')
    upm = font['head'].unitsPerEm
    if upm <= 0:
        raise ValueError('Invalid font units per em')
    cmap = dict(font.getBestCmap() or {})
    if not cmap:
        for table in font['cmap'].tables:
            if table.platformID == 3 and table.platEncID == 0:
                cmap.update(table.cmap)
    if any(not 0 <= cp <= 0x10FFFF or 0xD800 <= cp <= 0xDFFF for cp in cmap):
        raise ValueError('Cmap contains an invalid Unicode scalar')

    italic = {}
    extended = set()
    corners = {}
    has_math = 'MATH' in font
    if has_math:
        info = font['MATH'].table.MathGlyphInfo
        if info is not None:
            correction = info.MathItalicsCorrectionInfo
            if correction:
                italic = dict(zip(correction.Coverage.glyphs,
                                  [r.Value for r in correction.ItalicsCorrection]))
            if info.ExtendedShapeCoverage:
                extended = set(info.ExtendedShapeCoverage.glyphs)
            if info.MathKernInfo:
                kern = info.MathKernInfo
                for name, record in zip(kern.MathKernCoverage.glyphs, kern.MathKernInfoRecords):
                    values = {}
                    for attr, key in [('TopRightMathKern', 'top_right'),
                                      ('TopLeftMathKern', 'top_left'),
                                      ('BottomRightMathKern', 'bottom_right'),
                                      ('BottomLeftMathKern', 'bottom_left')]:
                        table = getattr(record, attr)
                        if table:
                            heights = [v.Value for v in table.CorrectionHeight or []]
                            offsets = [v.Value for v in table.KernValue]
                            if len(offsets) != len(heights) + 1:
                                raise ValueError('Invalid MATH corner-kern dimensions')
                            values[key] = dict(heights=heights, values=offsets)
                    corners[str(font.getGlyphID(name))] = values

    original_glyf = None
    locations = None
    glyph_set = None
    if 'glyf' in font and stored_glyf:
        if font.reader is None:
            raise ValueError('Stored TrueType bounds require an original font reader')
        original_glyf = font.reader['glyf']
        locations = font['loca'].locations
    elif 'glyf' not in font:
        glyph_set = font.getGlyphSet()

    cache = {}

    def record(name):
        if name in cache:
            return cache[name]
        gid = font.getGlyphID(name)
        if original_glyf is not None:
            start, end = locations[gid:gid + 2]
            if start == end:
                bounds = [0, 0, 0, 0]
            else:
                if start < 0 or end > len(original_glyf) or end - start < 10:
                    raise ValueError('Invalid TrueType glyph header extent')
                bounds = list(struct.unpack_from('>hhhhh', original_glyf, start))[1:]
        elif 'glyf' in font:
            glyph = font['glyf'][name]
            bounds = [getattr(glyph, attr, 0) for attr in ('xMin', 'yMin', 'xMax', 'yMax')]
        else:
            pen = BoundsPen(glyph_set)
            glyph_set[name].draw(pen)
            bounds = list(pen.bounds or (0, 0, 0, 0))
        if bounds[0] > bounds[2] or bounds[1] > bounds[3]:
            raise ValueError('Invalid glyph bounds')
        value = [gid, font['hmtx'][name][0], *bounds, italic.get(name, 0)]
        cache[name] = value
        return value

    glyphs = {str(cp): record(name) for cp, name in sorted(cmap.items())}
    scripts = {'1': {}, '2': {}}
    source = font if ssty_source is None else ssty_source
    if 'GSUB' in source:
        table = source['GSUB'].table
        features = table.FeatureList.FeatureRecord if table.FeatureList else []
        lookups = sorted({i for f in features if f.FeatureTag == 'ssty'
                          for i in f.Feature.LookupListIndex})
        for index in lookups:
            lookup = table.LookupList.Lookup[index]
            for subtable in lookup.SubTable:
                sub = subtable.ExtSubTable if lookup.LookupType == 7 else subtable
                mapping = getattr(sub, 'alternates', None)
                if mapping is None:
                    mapping = getattr(sub, 'mapping', None)
                if mapping is None:
                    raise ValueError('Unsupported ssty lookup; script geometry was not exported')
                for cp, name in sorted(cmap.items()):
                    if name not in mapping:
                        continue
                    choices = mapping[name]
                    if isinstance(choices, str):
                        choices = [choices]
                    if not choices:
                        raise ValueError('Empty ssty alternate list')
                    for level in (1, 2):
                        value = record(choices[min(level - 1, len(choices) - 1)])
                        prior = scripts[str(level)].get(str(cp))
                        if prior is not None and prior != value:
                            raise ValueError('Conflicting ssty mappings')
                        scripts[str(level)][str(cp)] = value
    return dict(face_key=face_key, units_per_em=upm, has_math=has_math,
                glyphs=glyphs, script_glyphs=scripts, corner_kerns=corners,
                extended_gids=[font.getGlyphID(name) for name in sorted(extended)])
