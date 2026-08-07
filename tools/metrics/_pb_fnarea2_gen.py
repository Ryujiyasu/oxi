"""Word footnote-AREA reservation for a NO-TYPE docGrid Latin doc (0018715b geometry).

reports__0018715b (blind50) loses 4 paragraphs because Oxi's fn-constrained body
bottom is too generous — by >=0.8pt on its p5 and >=10pt on its p6, i.e. NOT a
constant, so the per-page AREA height (separator + the notes that land on that
page) is the suspect, not a fixed separator term.

S596b derived "no-grid separator = one footnote line" on bunkacontract (a TRUE
no-docGrid doc, Yu Gothic 10pt). This doc is a NO-TYPE docGrid (linePitch=360 ->
S571-refine leaves grid_pitch None) with Calibri 11 body / 10pt FootnoteText, and
its split above/below the separator rule was never measured.

Differential design (the _pb_fnres_gen.py shape): variant A carries no footnote;
variants B1/B2/B3 put 1/2/3 footnote references on early paragraphs (which always
stay on page 1). Sweeping the bottom margin and reading the first paragraph of
page 2 gives, for the same boundary filler,

    flip_A(cbot) - flip_Bk(cbot) = Word's TOTAL area reservation R(k)

so R(1) isolates separator + one note, and R(k) - R(k-1) is the marginal note.

Geometry = 0018715b: A4, pgMar 1440 all round, header 708 footer 708, docGrid
linePitch=360 (NO w:type), docDefaults Calibri sz=22 spacing after=160 line=259,
FootnoteText sz=20 after=0 line=240, separator paragraph after=0 line=240.
NOTE the target has NO compatibilityMode setting (the S933b class) - the settings
part here mirrors that (present, but no compat element).

Usage:
  python _pb_fnarea2_gen.py gen [fine:LO:HI:STEP:VAR]
  python _pb_fnarea2_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_fnarea2")

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

# target has a settings part but NO compatibilityMode element (S933b class)
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings {W_NS}><w:zoom w:percent="100"/></w:settings>')

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          '</w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
          '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>'
          '<w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
          '</w:styles>')

# short enough to be ONE line in the footnote area (so R(k) grows by one note line)
FN_TEXT = 'Short note number {n} for the area sweep.'


def footnotes(k):
    notes = ''
    for i in range(k):
        notes += ('<w:footnote w:id="%d"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
                  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
                  '<w:r><w:t xml:space="preserve"> %s</w:t></w:r></w:p></w:footnote>'
                  % (2 + i, FN_TEXT.format(n=i + 1)))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:footnotes {W_NS}>'
            '<w:footnote w:type="separator" w:id="-1"><w:p>'
            '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:separator/></w:r></w:p></w:footnote>'
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p>'
            '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
            + notes + '</w:footnotes>')


def para(i, refs):
    r = ''.join('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                '<w:footnoteReference w:id="%d"/></w:r>' % rid for rid in refs)
    return f'<w:p><w:r><w:t>Item {i:02d} alpha beta gamma delta.</w:t></w:r>{r}</w:p>'


def build(bottom_tw, k, n=34):
    # refs go on paragraphs 2..(k+1) — always on page 1
    body = ''
    for i in range(n):
        refs = [2 + (i - 1)] if (1 <= i <= k) else []
        body += para(i + 1, refs)
    body += (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1440" w:right="1440" w:bottom="{bottom_tw}" '
             f'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
             f'<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


CASES = [(k, b) for k in (0, 1, 2, 3) for b in range(1440, 2361, 40)]


def gen(cases=None):
    os.makedirs(OUTDIR, exist_ok=True)
    for k, b in (cases or CASES):
        with zipfile.ZipFile(os.path.join(OUTDIR, f'fa_{k}_{b:04d}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/settings.xml', SETTINGS)
            z.writestr('word/footnotes.xml', footnotes(k))
            z.writestr('word/document.xml', build(b, k))
    print('generated', len(cases or CASES), '->', OUTDIR)


def measure(pat='fa_*'):
    import glob
    import win32com.client
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
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
                cbot = 841.9 - b / 20.0
                print(f'{base}: cbot={cbot:7.2f} first_p2={fp if fp else "-"}', flush=True)
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
            gen([(int(var), b) for b in range(int(lo), int(hi) + 1, int(step))])
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'fa_*')
