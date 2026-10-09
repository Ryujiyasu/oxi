# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

import gzip
import hashlib
import json
from pathlib import Path
import tempfile
import unittest
import zipfile
import subprocess
import sys

from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTCollection, TTFont
from fontTools.ttLib.tables._c_m_a_p import CmapSubtable

from catalog import corpus_audit, export, pack_portable


def synthetic_face(bold=False):
    builder=FontBuilder(1000,isTTF=True)
    glyphs=['.notdef','A','kana','supplementary']
    builder.setupGlyphOrder(glyphs)
    builder.setupCharacterMap({65:'A',0x3042:'kana',0x20000:'supplementary'})
    builder.setupGlyf({name:TTGlyphPen(None).glyph() for name in glyphs})
    builder.setupHorizontalMetrics({name:(700 if bold else 600,0) for name in glyphs})
    builder.setupHorizontalHeader(ascent=800,descent=-200)
    style='Bold' if bold else 'Regular'
    builder.setupNameTable(dict(familyName='Catalog Fixture',styleName=style,
                               fullName='Catalog Fixture '+style,psName='CatalogFixture-'+style))
    builder.font['name'].setName('テスト書体',1,3,1,0x411)
    builder.setupOS2(sTypoAscender=800,sTypoDescender=-200,usWinAscent=800,usWinDescent=200,
                    fsSelection=32 if bold else 64)
    builder.setupPost()
    builder.setupMaxp()
    builder.font['head'].macStyle=1 if bold else 0
    return builder.font


class CatalogTests(unittest.TestCase):
    def test_portable_os2_codepage_and_average_width_follow_the_measured_face(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);fonts=root/'fonts';fonts.mkdir()
            with synthetic_face() as font:
                font['OS/2'].ulCodePageRange1=(1<<17)|1
                font['OS/2'].ulCodePageRange2=0
                font['OS/2'].xAvgCharWidth=777
                font.save(fonts/'metadata.ttf')
            with TTFont(fonts/'metadata.ttf') as font:
                average=font['OS/2'].xAvgCharWidth
                upm=font['head'].unitsPerEm
                codepage=font['OS/2'].ulCodePageRange1
            catalog=export([fonts],root/'catalog')
            self.assertFalse(catalog['errors'])
            packed=pack_portable(catalog,root/'catalog',root/'portable')
            face=packed['faces'][0]
            self.assertEqual(face['codepage_range1'],codepage)
            self.assertEqual(face['average_width_em'],average/upm)
            payload=(root/'portable/font_catalog_metrics.gz').read_bytes()
            raw=json.loads(gzip.decompress(payload[face['offset']:face['offset']+face['length']]))
            self.assertEqual(raw['average_width'],average)
            self.assertEqual(hashlib.sha256(payload).hexdigest(),packed['metrics_sha256'])

    def test_old_os2_without_codepage_fields_does_not_invent_a_charset(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);fonts=root/'fonts';fonts.mkdir()
            with synthetic_face() as font:
                font['OS/2'].version=0
                font['OS/2'].xAvgCharWidth=0
                font.save(fonts/'old.ttf')
            catalog=export([fonts],root/'catalog')
            self.assertFalse(catalog['errors'])
            packed=pack_portable(catalog,root/'catalog',root/'portable')
            self.assertIsNone(packed['faces'][0]['codepage_range1'])
            self.assertIsNone(packed['faces'][0]['average_width_em'])

    def test_collection_faces_localized_names_full_cmap_and_determinism(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);fonts=root/'fonts';fonts.mkdir()
            collection=TTCollection();collection.fonts=[synthetic_face(),synthetic_face(True)]
            collection.save(fonts/'fixture.ttc')
            a=export([fonts],root/'a');b=export([fonts],root/'b')
            self.assertFalse(a['errors'])
            self.assertEqual(len(a['faces']),2)
            self.assertEqual(len(a['aliases']['テスト書体']),2)
            self.assertEqual([f['bold'] for f in a['faces']],[False,True])
            data=json.loads(gzip.decompress((root/'a/metrics.json.gz').read_bytes()))
            self.assertEqual([data[f['key']]['widths']['131072'] for f in a['faces']],[600,700])
            self.assertEqual((root/'a/catalog.json').read_bytes(),(root/'b/catalog.json').read_bytes())
            self.assertEqual(a['metrics_sha256'],b['metrics_sha256'])
            packed=pack_portable(a,root/'a',root/'portable')
            payload=(root/'portable/font_catalog_metrics.gz').read_bytes()
            for face in packed['faces']:
                start=face['offset'];end=start+face['length']
                self.assertEqual(json.loads(gzip.decompress(payload[start:end])),data[face['key']])
            self.assertNotIn(str(root),json.dumps(packed))

    def test_symbol_cmap_is_not_treated_as_missing_widths(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);fonts=root/'fonts';fonts.mkdir()
            font=synthetic_face()
            table=CmapSubtable.newSubtable(4)
            table.platformID=3;table.platEncID=0;table.language=0
            table.cmap={0xF041:'A'}
            font['cmap'].tables=[table]
            font.save(fonts/'symbol.ttf')
            result=export([fonts],root/'out')
            data=json.loads(gzip.decompress((root/'out/metrics.json.gz').read_bytes()))
            widths=data[result['faces'][0]['key']]['widths']
            self.assertEqual(widths,{'61505':600})

    def test_named_variable_instances_are_recorded_separately(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);fonts=root/'fonts';fonts.mkdir()
            builder=FontBuilder(font=synthetic_face())
            builder.setupFvar([('wght',100,400,900,'Weight')],[
                dict(location={'wght':400},stylename='Regular'),
                dict(location={'wght':700},stylename='Bold'),
            ])
            builder.setupGvar({glyph:[] for glyph in builder.font.getGlyphOrder()})
            builder.font.save(fonts/'variable.ttf')
            result=export([fonts],root/'out')
            self.assertFalse(result['errors'])
            instances=[f for f in result['faces'] if 'instance_coordinates' in f]
            self.assertEqual(len(instances),2)
            self.assertEqual([f['weight'] for f in instances],[400,700])
            self.assertEqual([f['bold'] for f in instances],[False,True])

    def test_duplicate_source_keeps_provenance_without_duplicate_faces(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);fonts=root/'fonts';fonts.mkdir()
            synthetic_face().save(fonts/'a.ttf')
            (fonts/'b.ttf').write_bytes((fonts/'a.ttf').read_bytes())
            result=export([fonts],root/'out')
            self.assertEqual(len(result['faces']),1)
            self.assertEqual(len(result['faces'][0]['sources']),2)

    def test_invalid_fonts_and_unresolved_corpus_names_are_explicit(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);fonts=root/'fonts';fonts.mkdir()
            (fonts/'bad.ttf').write_bytes(b'not a font')
            out=root/'out';result=export([fonts],out)
            self.assertEqual(len(result['errors']),1)
            doc=root/'fixture.docx'
            with zipfile.ZipFile(doc,'w') as z:
                z.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rFonts w:ascii="Missing"/></w:document>')
            report=corpus_audit([doc],result,out)
            self.assertEqual(report['documents'],1)
            self.assertEqual(report['rows'][0]['status'],'unresolved_name')
            lock=root/'~$fixture.docx';lock.write_bytes(b'Word owner record')
            report=corpus_audit([lock],result,out)
            self.assertEqual(report['documents'],0)
            self.assertFalse(report['errors'])
            self.assertEqual(len(report['skipped']),1)

    def test_required_font_check_has_a_nonzero_exit_for_missing_names(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);fonts=root/'fonts';fonts.mkdir()
            synthetic_face().save(fonts/'fixture.ttf')
            export([fonts],root/'catalog')
            command=[sys.executable,str(Path(__file__).with_name('catalog.py')),
                     '--reuse-catalog',str(root/'catalog/catalog.json'),
                     '--output',str(root/'audit'),'--require-font']
            found=subprocess.run(command+['Catalog Fixture'],capture_output=True)
            missing=subprocess.run(command+['Missing Fixture'],capture_output=True)
            self.assertEqual(found.returncode,0,found.stderr)
            self.assertEqual(missing.returncode,2,missing.stderr)
            postscript=subprocess.run(command+['CatalogFixture-Regular'],capture_output=True)
            self.assertEqual(postscript.returncode,2,postscript.stderr)


if __name__=='__main__':unittest.main()
