# -*- coding: utf-8 -*-
"""contextualSpacing below-gap suppression discriminator (the S861 counterexample).

S861 (shipped): ctx on the UPPER para removes the WHOLE below-gap incl. the
lower's own before (derived on educational__000555ad: upper = ctx NUMBERED
Normal item, lower = plain Normal, before=180 -> Word gap = 1 line).
COUNTEREXAMPLE (forms__00042714 wi179->180): upper = ctx Normal (NO numPr,
ind left=634), lower = plain Normal before=360 -> Word KEEPS the 18pt
(box 32.2 = line 14.2 + 18).

Hypothesis: the suppression requires the ctx para to be a LIST item (numPr).
Configs (single page, read gap = Info6(B) - Info6(A)):
  c1: A plain           , B before=360   -> control, expect line+18
  c2: A ctx             , B before=360   -> forms shape: KEEP?
  c3: A ctx + numPr     , B before=360   -> educational shape: SUPPRESS?
  c4: A ctx + numPr     , B numPr + before=360 (same list)
  c5: A ctx + ind left  , B before=360   -> forms exact (ind 634 + tabs)
  c6: A ctx             , B before=360 + ctx  (both ctx)
"""
import os, sys, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs3")
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
           '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/>'
          '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>'
          '</w:style></w:styles>')
NUMBERING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<w:numbering {W_NS}>'
             '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
             '<w:start w:val="1"/><w:numFmt w:val="decimal"/>'
             '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
             '<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl>'
             '</w:abstractNum>'
             '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
             '</w:numbering>')
R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
NUMPR = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'


def para(text, ppr):
    return (f'<w:p><w:pPr>{ppr}</w:pPr>'
            f'<w:r><w:rPr>{R}</w:rPr><w:t>{text}</w:t></w:r></w:p>')


CONFIGS = {
    'c1': (SP0,                                      f'<w:spacing w:before="360" w:after="0"/>'),
    'c2': (SP0 + '<w:contextualSpacing/>',           f'<w:spacing w:before="360" w:after="0"/>'),
    'c3': (NUMPR + SP0 + '<w:contextualSpacing/>',   f'<w:spacing w:before="360" w:after="0"/>'),
    'c4': (NUMPR + SP0 + '<w:contextualSpacing/>',   NUMPR + f'<w:spacing w:before="360" w:after="0"/>'),
    'c5': ('<w:tabs><w:tab w:val="left" w:pos="5040"/></w:tabs>'
           '<w:spacing w:after="0"/><w:ind w:left="634"/><w:contextualSpacing/>',
           '<w:tabs><w:tab w:val="left" w:pos="10080"/></w:tabs>'
           '<w:spacing w:before="360" w:after="0"/>'),
    'c6': (SP0 + '<w:contextualSpacing/>',
           f'<w:spacing w:before="360" w:after="0"/><w:contextualSpacing/>'),
}


def build(cfg):
    pa, pb = CONFIGS[cfg]
    body = (para('Filler zero line.', SP0)
            + para('AAA upper paragraph.', pa)
            + para('BBB lower paragraph.', pb)
            + para('CCC tail.', SP0)
            + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1440" w:right="1418" w:bottom="1440" w:left="1418" '
              'w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for cfg in CONFIGS:
        with zipfile.ZipFile(os.path.join(OUTDIR, f'cs_{cfg}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', build(cfg))
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/numbering.xml', NUMBERING)
    print(f'generated {len(CONFIGS)}')


def measure():
    import win32com.client
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    out = {}
    try:
        for fn in sorted(f for f in os.listdir(OUTDIR) if f.endswith('.docx')):
            d = word.Documents.Open(os.path.abspath(os.path.join(OUTDIR, fn)), ReadOnly=True)
            try:
                ys = []
                for i in range(1, d.Paragraphs.Count + 1):
                    rng = d.Paragraphs(i).Range
                    r0 = d.Range(rng.Start, rng.Start)
                    ys.append(r0.Information(6))
                gap_ab = ys[2] - ys[1]   # B.y - A.y  (A's box + applied before)
                out[fn[3:-5]] = {'ys': [round(y, 2) for y in ys], 'gapAB': round(gap_ab, 2)}
                print(f'  {fn}: gap A->B = {gap_ab:.2f}  (line 12.65 + before18 = 30.65 if KEPT; ~12.65 if SUPPRESSED)')
            finally:
                d.Close(False)
    finally:
        word.Quit()
    json.dump(out, open(os.path.join(OUTDIR, '_measure.json'), 'w'), indent=1)


if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == 'gen': gen()
    elif a and a[0] == 'measure': measure()
    else: print(__doc__)
