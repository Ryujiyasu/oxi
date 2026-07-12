"""Footnote REF-MARK line height derivation — font × size sweep.

The S804 footnote-area model left the ref-line growth underived (TNR10:
ref line renders 15.0 vs plain 11.5 = +3.5; Calibri10.5: +1.08). The ref
line = the note's first line containing the superscript footnote mark.
Word's fn area = sep_region + ref_line_h + (n-1)*plain + after, bottom-
anchored, so the ref-line height is recoverable from the note baselines
directly (b2 - b1 vs b3 - b2 does NOT give it — line pitch measures the
FOLLOWING line's height; instead the AREA TOP does): measure b1..b3 and
the separator rule y per config.

Configs: font {Times New Roman, Arial, Calibri} × sz {18,20,22,24} (9-12pt).
Body filler pushes the page full so the area is bottom-anchored.

Usage:
  python _fn_refline_gen.py gen
  python _fn_refline_gen.py measure
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_fn_refline")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
           '</Relationships>')


def styles(font, sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:styles {W_NS}>'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
            '<w:pPr/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
            f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz}"/></w:rPr></w:style>'
            '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>'
            '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
            '</w:styles>')

FN_TEXT = ('REFLINE first line of the footnote body long enough to wrap onto the '
           'second printed line and then the third printed line of guidance text '
           'for the pitch measurement of the plain footnote line grid xyz end')

FOOTNOTES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<w:footnotes {W_NS}>'
             '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
             '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
             '<w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
             '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
             f'<w:r><w:t xml:space="preserve"> {FN_TEXT}</w:t></w:r></w:p></w:footnote>'
             '</w:footnotes>')


def body_para(i, ref=False):
    r = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
    runs = f'<w:r><w:rPr>{r}</w:rPr><w:t>Body paragraph {i:02d} filler content sentence.</w:t></w:r>'
    if ref:
        runs += '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r>'
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{r}</w:rPr></w:pPr>{runs}</w:p>')


CASES = [(f, sz) for f in ('Times New Roman', 'Arial', 'Calibri')
         for sz in (18, 20, 22, 24)]


def name(c):
    return f"fnr_{c[0].replace(' ', '')[:6]}_s{c[1]}.docx"


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for c in CASES:
        font, sz = c
        body = ''.join(body_para(i + 1, ref=(i == 5)) for i in range(70))
        body += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1440" w:right="851" w:bottom="1440" '
                 'w:left="851" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')
        with zipfile.ZipFile(os.path.join(OUTDIR, name(c)), 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', doc)
            z.writestr('word/styles.xml', styles(font, sz))
            z.writestr('word/footnotes.xml', FOOTNOTES)
    print('generated', len(CASES), 'docs in', OUTDIR)


def measure():
    import win32com.client, fitz
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        for c in CASES:
            p = os.path.abspath(os.path.join(OUTDIR, name(c)))
            pdf = p[:-5] + '.pdf'
            doc = word.Documents.Open(p, ReadOnly=True)
            doc.ExportAsFixedFormat(pdf, 17)
            doc.Close(False)
            d = fitz.open(pdf)
            pg = d[0]
            rule = None
            for dr in pg.get_drawings():
                r = dr['rect']
                if r.height < 3 and r.width > 50 and r.y0 > 400:
                    rule = round(r.y0, 2)
            body_last = None
            fn = []
            for blk in pg.get_text('dict')['blocks']:
                if blk.get('type') != 0:
                    continue
                for ln in blk['lines']:
                    t = ''.join(s['text'] for s in ln['spans'])
                    if not t.strip():
                        continue
                    y = round(ln['spans'][0]['origin'][1], 2)
                    if 'Body paragraph' in t:
                        body_last = max(body_last or 0, y)
                    elif 'REFLINE' in t or 'printed line' in t or 'xyz end' in t or 'guidance' in t:
                        fn.append(y)
            fn.sort()
            font, sz = c
            pitches = [round(b - a, 2) for a, b in zip(fn, fn[1:])]
            print(f'{name(c)[:-5]}: body_last={body_last} rule={rule} fn_bases={fn} pitches={pitches}')
    finally:
        word.Quit()


if __name__ == '__main__':
    if sys.argv[1] == 'gen':
        gen()
    else:
        measure()
