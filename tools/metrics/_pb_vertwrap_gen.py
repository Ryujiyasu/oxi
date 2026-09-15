# -*- coding: utf-8 -*-
"""How does a square-wrap anchor cut the columns of a vertical (tbRl) section?

reports__2970ce67a2c26f7e (tbRl, 5 bands, linePitch 360, 86 empty body
paragraphs, 13 anchored shapes carrying all the visible text): Word lays the
body over 2 pages, Oxi over 1 with a blank page -- the vertical path places no
anchors at all, so nothing cuts its columns. Word's truth x-walk on p1 shows
band 1/3/4 holding no line, band 2 holding 9 lines of 36pt from x=381.75
leftwards (a right-side shape at page x 390.7 with distL 9), band 5 holding 18
lines of 18pt (a shape ending at y=722 leaves 47pt of the band).

Sheet: A4, margins 1440/1080, tbRl, docGrid lines 360, Normal 10.5pt Yu
Mincho (minorEastAsia), body = 60 empty paragraphs. One wps rectangle anchored
in paragraph 1, square wrap, positioned from the margin. Readout: per-paragraph
page / x / y via Information(3/5/6) on the collapsed start.

Arms (name: shape box in pt, margin-relative; bands):
  none       -- control, 5 bands
  right_full -- x 300..487 y 0..698 (a right-side strip through every band)
  left_full  -- x 0..187 y 0..698
  mid_full   -- x 150..300 y 0..698 (strips on both sides)
  top_R<r>   -- x 0..487, y 0..(band_h - r): band 1 keeps r pt of room;
                r in 5, 10, 12, 18, 20, 30, 40 -> the smallest room a line needs
  band1_x    -- x 100..400 y 0..122 (band 1 only; bands 2-5 untouched)
  one_band   -- same box, cols num=1 (a single band)
  dist0      -- right_full with distL=distR=0
  tb_full    -- right_full with wrapTopAndBottom (control for the wrap kind)
"""
import os, sys, json, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/vertwrap'); OUT.mkdir(parents=True, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
      'mc:Ignorable="wp14"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '</w:styles>')
EMU = 12700


def sect(ncols):
    cols = f'<w:cols w:num="{ncols}" w:space="424"/>' if ncols > 1 else '<w:cols w:space="424"/>'
    return ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="851" w:footer="992" w:gutter="0"/>'
            f'{cols}<w:textDirection w:val="tbRl"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')


def anchor(x, y, w, h, dist=9, wrap='square'):
    d = int(dist * EMU)
    wrapx = {'square': '<wp:wrapSquare wrapText="bothSides"/>', 'tb': '<wp:wrapTopAndBottom/>', 'none': '<wp:wrapNone/>'}[wrap]
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="{d}" distR="{d}" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="margin"><wp:posOffset>{int(x * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="margin"><wp:posOffset>{int(y * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(w * EMU)}" cy="{int(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrapx}'
            '<wp:docPr id="1" name="Rect 1"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp><wps:cNvSpPr/><wps:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
            '</wps:spPr><wps:bodyPr rtlCol="0" anchor="ctr"/></wps:wsp>'
            '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def document(shape, ncols, n=60, sz=21):
    paras = []
    ppr = f'<w:pPr><w:rPr><w:b/><w:sz w:val="{sz}"/></w:rPr></w:pPr>' if sz != 21 else ''
    for i in range(n):
        body = shape if (i == 0 and shape) else ''
        paras.append(f'<w:p>{ppr}{body}</w:p>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(paras)}{sect(ncols)}</w:body></w:document>'


def build(path, shape, ncols, sz=21):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/document.xml', document(shape, ncols, sz=sz))


BAND_H = (841.9 - 144 - 4 * 21.2) / 5  # 122.6
ARMS = {
    'none': (None, 5),
    'right_full': (anchor(300, 0, 187, 698), 5),
    'left_full': (anchor(0, 0, 187, 698), 5),
    'mid_full': (anchor(150, 0, 150, 698), 5),
    'band1_x': (anchor(100, 0, 300, 122), 5),
    'one_band': (anchor(100, 0, 300, 122), 1),
    'dist0': (anchor(300, 0, 187, 698, dist=0), 5),
    'tb_full': (anchor(300, 0, 187, 698, wrap='tb'), 5),
    'none_wrap': (anchor(300, 0, 187, 698, wrap='none'), 5),
    'band1_tb': (anchor(100, 0, 300, 122, wrap='tb'), 5),
    'one_band_tb': (anchor(100, 0, 300, 122, wrap='tb'), 1),
    'one_band_low': (anchor(100, 300, 300, 122), 1),
    'one_band_low_tb': (anchor(100, 300, 300, 122, wrap='tb'), 1),
    'right_partial': (anchor(300, 0, 187, 300), 5),
}
for r in (14, 18, 20, 24, 30, 40):
    ARMS[f'top_R{r:02d}_16pt'] = (anchor(0, 0, 487, BAND_H - r), 5, 32)
ARMS['none_16pt'] = (None, 5, 32)
for r in (5, 10, 12, 14, 15, 16, 17, 18, 20, 30, 40):
    ARMS[f'top_R{r:02d}'] = (anchor(0, 0, 487, BAND_H - r), 5)

only = set(sys.argv[1:])
import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try:
    app.Visible = False
except Exception:
    pass
results = {}
if (OUT / '_word.json').exists():
    results = json.loads((OUT / '_word.json').read_text(encoding='utf-8'))
try:
    for name, arm in ARMS.items():
        shape, ncols = arm[0], arm[1]
        sz = arm[2] if len(arm) > 2 else 21
        if only and name not in only:
            continue
        at = OUT / f'{name}.docx'
        build(at, shape, ncols, sz=sz)
        doc = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            rows = []
            for i in range(1, doc.Paragraphs.Count + 1):
                rng = doc.Paragraphs(i).Range
                c = doc.Range(rng.Start, rng.Start)
                rows.append((int(c.Information(3)), round(float(c.Information(5)), 2), round(float(c.Information(6)), 2)))
            shp = []
            for s in doc.Shapes:
                shp.append((round(s.Left, 2), round(s.Top, 2), round(s.Width, 2), round(s.Height, 2), s.WrapFormat.Type))
            pages = doc.ComputeStatistics(2)
        finally:
            doc.Close(False)
        results[name] = {'pages': pages, 'shapes': shp, 'rows': rows}
        # summarise: per (page, y) band -> count, x range
        from collections import OrderedDict
        bands = OrderedDict()
        for (pg, x, y) in rows:
            bands.setdefault((pg, y), []).append(x)
        summ = '; '.join(f'p{pg} y{y}: n={len(xs)} x {max(xs)}..{min(xs)}' for (pg, y), xs in bands.items())
        print(f'{name:12s} pages={pages} shapes={shp} | {summ}', flush=True)
finally:
    app.Quit()
(OUT / '_word.json').write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding='utf-8')
