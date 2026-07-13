"""Word footnote RESERVATION derivation — differential bottom sweep (uklocal geometry).

uklocal wp5 (natural flow): Oxi keeps «metadata...» at eff_bot 645.3, Word
pushes (eff < 636) — Oxi's fn reservation is ~10-23pt SHORT for the
no-type-grid + FootnoteText(basedOn Normal sb/sa 240/240, TNR 10) shape.

Differential design: variant A (no footnote) vs variant B (same fillers,
para 2 carries a footnoteReference; fn body = FootnoteText para that wraps
2 lines). Sweep the BOTTOM margin; readout = the first paragraph of page 2
(Word COM, collapsed start). For the same boundary filler k:
    flip_B(cbot) - flip_A(cbot) = Word's TOTAL fn reservation R
(the footer term cancels; the sep/spacing/line split can then be checked
against Oxi's footnote_sep_alloc + estimate_footnote_h).

Geometry = uklocal: A4, pgMar top=1440 right=851 bottom=SWEPT left=851,
docGrid linePitch=360 (no type), Normal = Arial 11 sb/sa 240/240
widowControl=0, compat-15 settings.

Usage:
  python _pb_fnres_gen.py gen [fine:LO:HI:STEP:VAR]
  python _pb_fnres_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_fnres")

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
           '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
           '</Relationships>')

SETTINGS15 = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              f'<w:settings {W_NS}>'
              '<w:compat><w:compatSetting w:name="compatibilityMode" '
              'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
              '</w:settings>')

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
          '<w:lang w:val="en-GB" w:eastAsia="en-GB" w:bidi="ar-SA"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="240" w:after="240"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:widowControl/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
          '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/>'
          '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
          '</w:styles>')

FN_TEXT = ('Please note that some of the frequently asked questions referring to '
           'the transparency code are still published on the departmental website '
           'and remain valid for reference purposes today.')

FOOTNOTES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<w:footnotes {W_NS}>'
             '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'
             '<w:r><w:separator/></w:r></w:p></w:footnote>'
             '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'
             '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
             '<w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
             '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
             f'<w:r><w:t xml:space="preserve"> {FN_TEXT}</w:t></w:r></w:p></w:footnote>'
             '</w:footnotes>')


def para(i, with_ref=False):
    ref = ('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
           '<w:footnoteReference w:id="2"/></w:r>') if with_ref else ''
    return (f'<w:p><w:r><w:t>Item {i:02d} alpha beta gamma.</w:t></w:r>{ref}</w:p>')


def build(bottom_tw, variant, n=40):
    body = ''.join(para(i + 1, with_ref=(variant == 'B' and i == 1)) for i in range(n))
    body += (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1440" w:right="851" w:bottom="{bottom_tw}" '
             f'w:left="851" w:header="709" w:footer="709" w:gutter="0"/>'
             f'<w:docGrid w:linePitch="360"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


CASES = [(v, b) for v in ('A', 'B') for b in range(1300, 1921, 20)]


def gen(cases=None):
    os.makedirs(OUTDIR, exist_ok=True)
    for v, b in (cases or CASES):
        with zipfile.ZipFile(os.path.join(OUTDIR, f'fnr_{v}_{b:04d}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/settings.xml', SETTINGS15)
            z.writestr('word/footnotes.xml', FOOTNOTES)
            z.writestr('word/document.xml', build(b, v))
    print('generated', len(cases or CASES))


def measure(pat='fnr_*'):
    import glob
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    out = {}
    try:
        for f in sorted(glob.glob(os.path.join(OUTDIR, pat + '.docx'))):
            doc = word.Documents.Open(os.path.abspath(f), ReadOnly=True)
            try:
                fp = None
                for i in range(1, doc.Paragraphs.Count + 1):
                    rng = doc.Paragraphs(i).Range
                    if doc.Range(rng.Start, rng.Start).Information(3) >= 2:
                        fp = i
                        break
                base = os.path.basename(f)[:-5]
                b = int(base.rsplit('_', 1)[-1])
                print(f'{base}: cbot={841.9 - b / 20.0:.1f} first_p2=Item{fp:02d}' if fp
                      else f'{base}: 1page', flush=True)
                out[base] = fp
            finally:
                doc.Close(False)
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        spec = sys.argv[2] if len(sys.argv) > 2 else 'coarse'
        if spec == 'coarse':
            gen()
        else:
            _, lo, hi, step, var = spec.split(':')
            gen([(var, b) for b in range(int(lo), int(hi) + 1, int(step))])
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'fnr_*')
