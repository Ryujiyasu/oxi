# -*- coding: utf-8 -*-
"""Does a paragraph's w:leftChars="0" override a numbering level's w:left?

policies__0568bf409f366762 para 12: pPr `<w:ind w:leftChars="0"/>` (no w:left),
style List Paragraph `leftChars=400 left=840`, numbering lvl0 `left=1260
hanging=420`. Word resolves LeftIndent 63 / FirstLineIndent -21 -- the numbering
level's indent -- and wraps at 36/35 chars (3 lines); Oxi took leftChars=0 as
"zero", wrapped at 40 chars (2 lines) and fitted one paragraph more on p1.

Readout: Paragraph.LeftIndent / FirstLineIndent through Word COM (no render).
Measured 2026-09-15 (first 10 arms):
  i    lc0 + num + list style        -> 63.00 / -21   (numbering left wins)
  ii   lc0 + left=0 + num            ->  0.00 / -21   (an explicit left=0 wins)
  iii  lc100 + num                   -> 31.50 / -21   (= 1 char 10.5 + hanging 21)
  iv   num + list style, no para ind -> 63.00 / -21   (numbering beats style)
  v    lc0 + list style, no num      -> 42.00 /   0   (style left=840 stays)
  vi   lc0 + style left=600, no num  -> 30.00 /   0   (style absolute left stays)
  vii  lc0 + num, Normal style       -> 63.00 / -21
  viii left=300 + num                -> 63.00 / -21   (! a bare direct left loses)
  ix   lc0 + hc0 + num               -> 63.00 / -21
  x    firstLineChars=0 + num        -> 63.00 / -21
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/leftchars0'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
STY = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
       '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
       '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
       '<w:style w:type="paragraph" w:styleId="a3"><w:name w:val="List Paragraph"/><w:basedOn w:val="a"/><w:pPr><w:ind w:leftChars="400" w:left="840"/></w:pPr></w:style>'
       '<w:style w:type="paragraph" w:styleId="ab"><w:name w:val="Abs Style"/><w:basedOn w:val="a"/><w:pPr><w:ind w:left="600"/></w:pPr></w:style>'
       '<w:style w:type="paragraph" w:styleId="ac"><w:name w:val="Chars Style"/><w:basedOn w:val="a"/><w:pPr><w:ind w:leftChars="300" w:left="200"/></w:pPr></w:style></w:styles>')
NUM = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
       '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="・"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1260" w:hanging="420"/></w:pPr></w:lvl></w:abstractNum>'
       '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'
TEXT = 'コロナ禍の状況も踏まえ、相談先を待っている多くの方の期待に応え寄り添い、その当事者の皆さんの思いや願い、要求を実現する取り組みにつなげていきます。'
NUMPR = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
ARMS = {
  'i_lc0_num_liststyle':   f'<w:pStyle w:val="a3"/>{NUMPR}<w:ind w:leftChars="0"/>',
  'ii_lc0_left0_num':      f'<w:pStyle w:val="a3"/>{NUMPR}<w:ind w:leftChars="0" w:left="0"/>',
  'iii_lc100_num':         f'<w:pStyle w:val="a3"/>{NUMPR}<w:ind w:leftChars="100"/>',
  'iv_num_liststyle_only': f'<w:pStyle w:val="a3"/>{NUMPR}',
  'v_lc0_liststyle_nonum': '<w:pStyle w:val="a3"/><w:ind w:leftChars="0"/>',
  'vi_lc0_abs_style':      '<w:pStyle w:val="ab"/><w:ind w:leftChars="0"/>',
  'vii_lc0_num_normal':    f'{NUMPR}<w:ind w:leftChars="0"/>',
  'viii_left300_num':      f'<w:pStyle w:val="a3"/>{NUMPR}<w:ind w:left="300"/>',
  'ix_lc0_hc0_num':        f'<w:pStyle w:val="a3"/>{NUMPR}<w:ind w:leftChars="0" w:hangingChars="0"/>',
  'x_fl0_num':             f'<w:pStyle w:val="a3"/>{NUMPR}<w:ind w:firstLineChars="0"/>',
  # second batch: the bare-direct-left anomaly and chars-style fallbacks
  'xi_left300_hang420_num': f'<w:pStyle w:val="a3"/>{NUMPR}<w:ind w:left="300" w:hanging="420"/>',
  'xii_left2000_num':      f'<w:pStyle w:val="a3"/>{NUMPR}<w:ind w:left="2000"/>',
  'xiii_left300_normal_num': f'{NUMPR}<w:ind w:left="300"/>',
  'xiv_lc0_left300_num':   f'<w:pStyle w:val="a3"/>{NUMPR}<w:ind w:leftChars="0" w:left="300"/>',
  'xv_lc0_charsstyle_nonum': '<w:pStyle w:val="ac"/><w:ind w:leftChars="0"/>',
  'xvi_charsstyle_nonum':  '<w:pStyle w:val="ac"/>',
  'xvii_lc0_hang420_num':  f'<w:pStyle w:val="a3"/>{NUMPR}<w:ind w:leftChars="0" w:hanging="420"/>',
  'xviii_left300_nonum':   '<w:ind w:left="300"/>',
}
only = set(sys.argv[1:])
import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for name, ppr in ARMS.items():
        if only and name not in only:
            continue
        doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>{TEXT}</w:t></w:r></w:p>{SECT}</w:body></w:document>'
        at = OUT / f'{name}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/styles.xml', STY); z.writestr('word/numbering.xml', NUM); z.writestr('word/document.xml', doc)
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            p = d.Paragraphs(2)
            print(f'{name:26s} LeftIndent={p.LeftIndent:6.2f} FirstLine={p.FirstLineIndent:6.2f} lines={p.Range.ComputeStatistics(1)} list={p.Range.ListFormat.ListString!r}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
