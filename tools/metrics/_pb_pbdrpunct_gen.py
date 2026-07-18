"""Faithful probe: does Word (compat14, TNR11, BoxPara w/ pBdr) hang the
paragraph-final '.' past the content-right, or wrap?

Real doc truth (002c1ffa p73 (b) para): 'day' fits with 1.45pt to spare,
'day.' would end 1.28pt past content-right 474.85 -> Word WRAPS.
S809 (probe-derived, Arial legacy) predicts KEEP (full '.' hang).

Variants:
  faith    : styles.xml + settings.xml transplanted verbatim, BoxPara para
  nobdr    : same but pBdr stripped from BoxText
  arial    : faith but runs forced Arial
  notab    : faith but plain first-line (no tabs/marker), same ind
  plain    : bare para (no style), TNR 11, same ind/hanging, no pBdr
Each swept over right margin to find the 1->2 line flip; the flip position
vs (word) / (word+.) width identifies hang vs no-hang per variant.
"""
import os, sys, zipfile, shutil, re

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_pbdrpunct")
SRC = r"c:\Users\ryuji\oxi-main\pipeline_data\docx_corpus\en\technical\002c1ffa65f3a566.docx"

z = zipfile.ZipFile(SRC)
STYLES = z.read('word/styles.xml').decode('utf-8')
# Fresh minimal settings carrying the real doc's layout-relevant values
# (the verbatim transplant dies on attachedTemplate r:id / linkStyles).
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:defaultTabStop w:val="720"/>'
    '<w:noPunctuationKerning/>'
    '<w:characterSpacingControl w:val="doNotCompress"/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
    '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '<w:compatSetting w:name="enableOpenTypeFeatures" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '<w:compatSetting w:name="doNotFlipMirrorIndents" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '</w:compat></w:settings>')

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
           '</Relationships>')

ARIAL = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>'

def para(variant):
    if variant in ('faith', 'nobdr', 'arial'):
        rpr = ARIAL if variant == 'arial' else ''
        return ('<w:p><w:pPr><w:pStyle w:val="BoxPara"/>' +
                (f'<w:rPr>{ARIAL[7:-8]}</w:rPr>' if variant == 'arial' else '') +
                '</w:pPr>'
                f'<w:r>{rpr}<w:tab/><w:t>(b)</w:t></w:r>'
                f'<w:r>{rpr}<w:tab/><w:t xml:space="preserve">the parent’s multi</w:t></w:r>'
                f'<w:r>{rpr}<w:noBreakHyphen/></w:r>'
                f'<w:r>{rpr}<w:t>case cap for the child for the day.</w:t></w:r></w:p>')
    if variant == 'notab':
        return ('<w:p><w:pPr><w:pStyle w:val="BoxPara"/>'
                '<w:ind w:left="2552" w:firstLine="0"/></w:pPr>'
                '<w:r><w:t xml:space="preserve">the parent’s multi</w:t></w:r>'
                '<w:r><w:noBreakHyphen/></w:r>'
                '<w:r><w:t>case cap for the child for the day.</w:t></w:r></w:p>')
    if variant == 'plain':
        return ('<w:p><w:pPr><w:ind w:left="2552"/>'
                '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                '<w:r><w:t xml:space="preserve">the parent’s multi</w:t></w:r>'
                '<w:r><w:noBreakHyphen/></w:r>'
                '<w:r><w:t>case cap for the child for the day.</w:t></w:r></w:p>')
    raise ValueError(variant)

def build(variant, right_tw, path):
    styles = STYLES
    if variant == 'nobdr':
        styles = re.sub(r'<w:pBdr>.*?</w:pBdr>', '', styles, flags=re.S)
    body = para(variant)
    body += (f'<w:sectPr><w:pgSz w:w="11907" w:h="16839"/>'
             f'<w:pgMar w:top="1418" w:right="{right_tw}" w:bottom="1418" '
             f'w:left="2410" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as o:
        o.writestr('[Content_Types].xml', CT)
        o.writestr('_rels/.rels', RELS)
        o.writestr('word/_rels/document.xml.rels', DOCRELS)
        o.writestr('word/document.xml', doc)
        o.writestr('word/styles.xml', styles)
        o.writestr('word/settings.xml', SETTINGS)

# Sweep: right margin around 2410 (real). At 2410 the '.' crosses by 1.28pt.
# 'day.' = 18.6pt = 373tw; 'day' = 318tw; '.' = 55tw (2.75pt).
# If NO hang: flip (1->2 lines) at right margin where full 'day.' stops
#   fitting: need <= 1.28pt more room -> flip between 2384 (2410-26tw) and 2410.
# If FULL hang: flip when 'day' stops fitting: 'day' has 28tw spare at 2410 ->
#   flip between 2410 and 2438+.
# Sweep 2330..2480 step 10tw (0.5pt) covers both with margin.
CASES = []
for v in ('faith', 'nobdr', 'arial', 'notab', 'plain'):
    for r in range(2330, 2481, 10):
        CASES.append((v, r))

def gen():
    os.makedirs(OUT, exist_ok=True)
    for v, r in CASES:
        build(v, r, os.path.join(OUT, f'pp_{v}_{r}.docx'))
    print('generated', len(CASES))

def measure():
    import win32com.client, glob
    import fitz
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for v, r in CASES:
            p = os.path.join(OUT, f'pp_{v}_{r}.docx')
            pdf = p.replace('.docx', '.pdf')
            if not os.path.exists(pdf):
                doc = word.Documents.Open(p, ReadOnly=True)
                doc.ExportAsFixedFormat(pdf, 17)
                doc.Close(False)
            d = fitz.open(pdf)
            pg = d[0]
            lines = []
            for b in pg.get_text('dict')['blocks']:
                if b['type'] != 0: continue
                for l in b['lines']:
                    t = ''.join(s['text'] for s in l['spans']).strip()
                    if t: lines.append((round(l['bbox'][1], 2), t, round(l['bbox'][2], 1)))
            d.close()
            # count distinct y of text lines
            ys = sorted(set(y for y, t, x1 in lines))
            res[(v, r)] = (len(ys), lines)
    finally:
        word.Quit()
    import json
    with open(os.path.join(OUT, 'result.json'), 'w', encoding='utf-8') as f:
        json.dump({f'{v}_{r}': [n, lines] for (v, r), (n, lines) in res.items()}, f, ensure_ascii=False)
    # report flips
    for v in ('faith', 'nobdr', 'arial', 'notab', 'plain'):
        seq = [(r, res[(v, r)][0]) for _, r in [(v, r) for vv, r in CASES if vv == v]]
        seq.sort()
        flip = None
        for (r1, n1), (r2, n2) in zip(seq, seq[1:]):
            if n1 != n2:
                flip = (r1, n1, r2, n2)
        print(v, 'lines@2410:', res[(v, 2410)][0], 'flips:',
              [(r1, n1, r2, n2) for (r1, n1), (r2, n2) in zip(seq, seq[1:]) if n1 != n2])

if __name__ == '__main__':
    if sys.argv[1] == 'gen':
        gen()
    else:
        measure()
