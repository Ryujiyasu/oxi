# -*- coding: utf-8 -*-
"""STYLE-level HTML autospacing: the empty-paragraph adjacency rule.

legal__00081e80 (Metadata style, before=100 auto=1 after=100 auto=1):
Word applies ~13.95/gap between TEXT paragraphs. correspondence__00054c43
(NormalWeb, IDENTICAL style shape): the gaps AROUND an EMPTY NormalWeb
paragraph render at the explicit 5pt, not 13.75 (S895's first default-ON
flipped it). Hypothesis: HTML empty-block margin semantics — auto margins
do not materialize on any gap ADJACENT to an empty paragraph.

Configs (Normal = no spacing; Web = before=100 beforeAutospacing=1
after=100 afterAutospacing=1, basedOn Normal; TNR 12 everywhere):
  cTT   : Wa Wb Wc            (text-text gaps — expect ~line+13.75)
  cTET  : Wa WE Wb            (text-empty-text — the 00054c43 shape)
  cTE2T : Wa WE WE Wb         (double empty)
  cTNT  : Wa NE Wb            (Normal-style EMPTY between Web texts)
  cDT   : Wa Da Wb            (Da = Web text with DIRECT before=200 after=200
                               — the per-side override control)
Read: COM Information(6) per paragraph start (0.75-quantized; 5-vs-14
discrimination is far above the noise).

Usage:
  python _pb_autosp_gen.py gen
  python _pb_autosp_gen.py measure
"""
import os, sys, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_autosp")

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Web"><w:name w:val="Normal (Web)"/>'
          '<w:basedOn w:val="Normal"/>'
          '<w:pPr><w:spacing w:before="100" w:beforeAutospacing="1" '
          'w:after="100" w:afterAutospacing="1"/></w:pPr></w:style>'
          '</w:styles>')


def wtext(tag, style='Web', spacing=''):
    return (f'<w:p><w:pPr><w:pStyle w:val="{style}"/>{spacing}</w:pPr>'
            f'<w:r><w:t>{tag} lorem ipsum dolor.</w:t></w:r></w:p>')


def wempty(style='Web'):
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr></w:p>'


CONFIGS = {
    'cTT': wtext('A1') + wtext('B2') + wtext('C3'),
    'cTET': wtext('A1') + wempty() + wtext('B2'),
    'cTE2T': wtext('A1') + wempty() + wempty() + wtext('B2'),
    'cTNT': wtext('A1') + wempty('Normal') + wtext('B2'),
    'cDT': (wtext('A1')
            + wtext('D2', spacing='<w:spacing w:before="200" w:after="200"/>')
            + wtext('B2')),
}


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for cfg, body in CONFIGS.items():
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               f'<w:document {W_NS}><w:body>{body}'
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
               '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
               'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>'
               '</w:body></w:document>')
        with zipfile.ZipFile(os.path.join(OUTDIR, f'pas_{cfg}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', doc)
            z.writestr('word/styles.xml', STYLES)
    print('generated', len(CONFIGS), 'docs in', os.path.abspath(OUTDIR))


def measure():
    import glob
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for f in sorted(glob.glob(os.path.join(OUTDIR, 'pas_*.docx'))):
            d = word.Documents.Open(os.path.abspath(f), ReadOnly=True)
            ys = []
            try:
                for i in range(1, d.Paragraphs.Count + 1):
                    r = d.Paragraphs(i).Range
                    rs = d.Range(r.Start, r.Start)
                    t = ''.join(ch for ch in r.Text if ch.isprintable()).strip()
                    ys.append((round(rs.Information(6), 2), t[:10]))
            finally:
                d.Close(False)
            base = os.path.basename(f)[:-5]
            gaps = [round(b[0] - a[0], 2) for a, b in zip(ys, ys[1:])]
            res[base] = {'ys': ys, 'gaps': gaps}
            print(f'{base}: ys={[y for y, _ in ys]} gaps={gaps}')
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, '_measure.json'), 'w'), indent=1)


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        gen()
    else:
        measure()
