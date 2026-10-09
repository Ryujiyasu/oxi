# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

import gzip
import hashlib
import json
from pathlib import Path
import tempfile
import unittest

from fontTools.feaLib.builder import addOpenTypeFeaturesFromString
from fontTools.fontBuilder import FontBuilder
from fontTools.otlLib.builder import buildMathTable
from fontTools.pens.t2CharStringPen import T2CharStringPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont
from fontTools.ttLib.tables.TupleVariation import TupleVariation
from fontTools.ttLib.tables._c_m_a_p import CmapSubtable

from catalog import export, pack_portable
from font_geometry import face_geometry


def fixture_font(*, cff=False, math=False, variable=False):
    builder = FontBuilder(1000, isTTF=not cff)
    order = ['.notdef', 'space', 'A', 'A.script', 'A.script2', 'supplementary']
    builder.setupGlyphOrder(order)
    builder.setupCharacterMap({32: 'space', 65: 'A', 0x20000: 'supplementary'})
    glyphs = {}
    for index, name in enumerate(order):
        pen = T2CharStringPen(600, None) if cff else TTGlyphPen(None)
        if name not in ['.notdef', 'space']:
            pen.moveTo((10, -20))
            pen.lineTo((110 + index * 10, -20))
            pen.lineTo((110 + index * 10, 200 + index * 10))
            pen.lineTo((10, 200 + index * 10))
            pen.closePath()
        glyphs[name] = pen.getCharString() if cff else pen.glyph()
    if cff:
        builder.setupCFF('GeometryFixture', dict(FullName='Geometry Fixture',
                         FamilyName='Geometry Fixture', Weight='Regular'), glyphs, {})
    else:
        builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics({name: (600 - i * 20, 10) for i, name in enumerate(order)})
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    builder.setupNameTable(dict(familyName='Geometry Fixture', styleName='Regular',
                               fullName='Geometry Fixture Regular', psName='GeometryFixture'))
    builder.setupOS2(sTypoAscender=800, sTypoDescender=-200,
                    usWinAscent=800, usWinDescent=200, fsSelection=64)
    builder.setupPost()
    builder.setupMaxp()
    if math:
        buildMathTable(builder.font, italicsCorrections={'A': 30, 'A.script': 40},
                       extendedShapes={'A.script2'},
                       mathKerns={'A': {'TopRight': ([200], [-20, -10])}})
        addOpenTypeFeaturesFromString(builder.font,
            'feature ssty { sub A from [A.script A.script2]; } ssty;')
    if variable:
        builder.setupFvar([('wght', 100, 400, 900, 'Weight')], [
            dict(location={'wght': 400}, stylename='Regular'),
            dict(location={'wght': 700}, stylename='Bold')])
        variations = {}
        for name, glyph in builder.font['glyf'].glyphs.items():
            count = len(glyph.coordinates) if glyph.numberOfContours > 0 else 0
            deltas = [(100, 0)] * count + [(0, 0)] * 4
            variations[name] = [TupleVariation({'wght': (0, 1, 1)}, deltas)]
        builder.setupGvar(variations)
    return builder.font


class GlyphGeometryTests(unittest.TestCase):
    def test_symbol_geometry_preserves_actual_codepoints_without_ascii_aliases(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / 'symbol.ttf'
            with fixture_font() as font:
                table = CmapSubtable.newSubtable(4)
                table.platformID = 3
                table.platEncID = 0
                table.language = 0
                table.cmap = {0xF041: 'A'}
                font['cmap'].tables = [table]
                font.save(path)
            result = export([path], root / 'catalog')
            self.assertFalse(result['errors'])
            data = json.loads(gzip.decompress((root / 'catalog/metrics.json.gz').read_bytes()))
            raw = data[result['faces'][0]['key']]
            self.assertEqual(raw['glyph_geometry']['glyphs'],
                             {'61505': [2, 560, 10, -20, 130, 220, 0]})
            self.assertEqual(raw['widths'], {'61505': 560})
            self.assertNotIn('65', raw['glyph_geometry']['glyphs'])

    def test_math_ssty_and_empty_glyphs_survive_the_portable_roundtrip(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / 'fixture.ttf'
            with fixture_font(math=True) as font:
                font.save(path)
            result = export([path], root / 'catalog')
            self.assertFalse(result['errors'])
            packed = pack_portable(result, root / 'catalog', root / 'portable')
            payload = (root / 'portable/font_catalog_metrics.gz').read_bytes()
            face = packed['faces'][0]
            raw = json.loads(gzip.decompress(payload[face['offset']:face['offset'] + face['length']]))
            geometry = raw['glyph_geometry']
            self.assertEqual(geometry['face_key'], face['key'])
            self.assertEqual(geometry['glyphs']['65'], [2, 560, 10, -20, 130, 220, 30])
            self.assertEqual(geometry['glyphs']['32'], [1, 580, 0, 0, 0, 0, 0])
            self.assertEqual(geometry['glyphs']['131072'][0], 5)
            self.assertEqual(geometry['script_glyphs']['1']['65'], [3, 540, 10, -20, 140, 230, 40])
            self.assertEqual(geometry['script_glyphs']['2']['65'], [4, 520, 10, -20, 150, 240, 0])
            self.assertEqual(geometry['corner_kerns']['2']['top_right'],
                             dict(heights=[200], values=[-20, -10]))
            self.assertEqual(geometry['extended_gids'], [4])
            self.assertTrue(geometry['has_math'])
            self.assertEqual(hashlib.sha256(payload).hexdigest(), packed['metrics_sha256'])
            self.assertNotIn(str(root), json.dumps(raw))
            self.assertNotIn('outlines', json.dumps(raw))

    def test_cff_geometry_is_extracted_without_exporting_charstrings(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / 'fixture.otf'
            with fixture_font(cff=True) as font:
                font.save(path)
            result = export([path], root / 'catalog')
            self.assertFalse(result['errors'])
            data = json.loads(gzip.decompress((root / 'catalog/metrics.json.gz').read_bytes()))
            raw = data[result['faces'][0]['key']]
            self.assertEqual(raw['glyph_geometry']['glyphs']['65'], [2, 560, 10, -20, 130, 220, 0])
            self.assertFalse(raw['glyph_geometry']['has_math'])
            self.assertNotIn('CharStrings', json.dumps(raw))

    def test_named_variable_instances_use_changed_bounds_and_retained_ssty(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / 'variable.ttf'
            with fixture_font(math=True, variable=True) as font:
                font.save(path)
            result = export([path], root / 'catalog')
            self.assertFalse(result['errors'])
            data = json.loads(gzip.decompress((root / 'catalog/metrics.json.gz').read_bytes()))
            faces = {f['weight']: f for f in result['faces'] if 'instance_coordinates' in f}
            regular = data[faces[400]['key']]['glyph_geometry']
            bold = data[faces[700]['key']]['glyph_geometry']
            self.assertEqual(regular['glyphs']['65'][2:6], [10, -20, 130, 220])
            self.assertEqual(bold['glyphs']['65'][2:6], [70, -20, 190, 220])
            self.assertEqual(bold['script_glyphs']['1']['65'][2:6], [70, -20, 200, 230])
            self.assertEqual(bold['script_glyphs']['2']['65'][2:6], [70, -20, 210, 240])
            self.assertEqual(bold['face_key'], faces[700]['key'])

    def test_original_truetype_keeps_header_bounds_instead_of_tight_curve_bounds(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / 'curve.ttf'
            with fixture_font() as font:
                pen = TTGlyphPen(None)
                pen.moveTo((0, 0))
                pen.qCurveTo((50, 200), (100, 0))
                pen.closePath()
                font['glyf'].glyphs['A'] = pen.glyph()
                font.save(path)
            with TTFont(path, lazy=True) as font:
                geometry = face_geometry(font, 'synthetic:0')
            self.assertEqual(geometry['glyphs']['65'][2:6], [0, 0, 100, 200])

    def test_unsupported_ssty_is_an_explicit_error_instead_of_partial_success(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / 'unsupported.ttf'
            with fixture_font() as font:
                addOpenTypeFeaturesFromString(font,
                    'feature ssty { sub A A by A.script; } ssty;')
                font.save(path)
            result = export([path], root / 'catalog')
            self.assertEqual(len(result['errors']), 1)
            self.assertIn('Unsupported ssty lookup', result['errors'][0]['error'])
            self.assertEqual(result['faces'], [])


if __name__ == '__main__':
    unittest.main()
