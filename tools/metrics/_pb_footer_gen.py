"""Footer-reservation derivation — controlled bottom sweep WITH footers.

The _pb_latin sweep pinned the body page-bottom rule (keep iff line_top +
full hhea line <= cbot, == S779). uklocalspending shows Word's cbot with a
FOOTER = 719.81 (= 841.9 - footer_dist 35.45 - footer_stack 86.65: the
stack includes the FIRST para's before AND the LAST para's after with
exact/hhea line heights); Oxi reserves ~4pt less (empty footer paras
ignore line=exact + use floor-10tw + the S780 descent relief). This sweep
pins the stack model per footer config.

Configs:
  fU: uklocalspending replica — 3 paras (Footer-style PAGE, Footer-style
      empty, Normal empty); Footer basedOn Normal + line=260 exact;
      Normal Arial 11 before/after=240.
  fS: single Footer-style PAGE para (same styles).
  fP: single para with pBdr top + TNR12 (nyserda-like, Normal spacing 0).
  fE: 2 EMPTY Normal paras (spacing 240/240) — empty-line rule.

Body: 80 single-line spacing-0 Arial 11 paragraphs (Item i ...), A4,
top=1440, bottom SWEPT. cbot flip per config solves the stack.

Usage:
  python _pb_footer_gen.py gen [fine:LO:HI:STEP:CONFIG]
  python _pb_footer_gen.py measure
"""
import os, sys, json, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_footer")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
           '</Relationships>')

# uklocalspending styles: Normal Arial 11 before/after=240 widowControl=0;
# Footer basedOn Normal + tabs + line=260 exact.
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="240" w:after="240"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs>'
          '<w:spacing w:line="260" w:lineRule="exact"/></w:pPr></w:style>'
          '</w:styles>')

PAGE_FLD = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>1</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')

# nyserda-shape styles: Normal NO spacing; Footer basedOn Normal + pBdr top
# + before=240 + jc center, TNR12 body. Discriminates border/before models.
STYLES_NY_TMPL = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:styles {W_NS}>'
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    '<w:pPr/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>'
    '<w:pPr>{BORDER}<w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs>{SPACING}<w:jc w:val="center"/></w:pPr></w:style>'
    '</w:styles>')

BORDER_XML = '<w:pBdr><w:top w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr>'
SPACING_XML = '<w:spacing w:before="240"/>'

NY_FOOTER = ('<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>'
             f'{PAGE_FLD}</w:p>'
             '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>')

NY_CONFIGS = {
    'fN':  (BORDER_XML, SPACING_XML),   # border + before (nyserda replica)
    'fN2': ('', SPACING_XML),           # before only
    'fN3': (BORDER_XML, ''),            # border only
}

FOOTERS = {
    'fU': ('<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
           f'{PAGE_FLD}</w:p>'
           '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>'
           '<w:p/>'),
    'fS': ('<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
           f'{PAGE_FLD}</w:p>'),
    'fP': ('<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr>'
           '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
           '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:pPr>'
           '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
           f'<w:t>Page </w:t></w:r>{PAGE_FLD}</w:p>'),
    'fE': '<w:p/><w:p/>',
}


def body_para(i):
    r = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>Item {i} alpha beta gamma delta.</w:t></w:r></w:p>')


def ny_body_para(i):
    r = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>'
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>Item {i} alpha beta gamma delta.</w:t></w:r></w:p>')


def build(bottom_tw, cfg):
    if cfg in NY_CONFIGS:
        # nyserda shape: Letter page, TNR12 body, footer dist 720tw.
        body = ''.join(ny_body_para(i + 1) for i in range(80))
        body += (f'<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
                 f'<w:pgSz w:w="12240" w:h="15840"/>'
                 f'<w:pgMar w:top="1440" w:right="1800" w:bottom="{bottom_tw}" '
                 f'w:left="1800" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')
        ftr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               f'<w:ftr {W_NS}>{NY_FOOTER}</w:ftr>')
        b, s = NY_CONFIGS[cfg]
        styles = STYLES_NY_TMPL.replace('{BORDER}', b).replace('{SPACING}', s)
        return doc, ftr, styles
    body = ''.join(body_para(i + 1) for i in range(80))
    body += (f'<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
             f'<w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1440" w:right="851" w:bottom="{bottom_tw}" '
             f'w:left="851" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')
    ftr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:ftr {W_NS}>{FOOTERS[cfg]}</w:ftr>')
    return doc, ftr, STYLES


def gen(cases, cfg):
    os.makedirs(OUTDIR, exist_ok=True)
    for b in cases:
        p = os.path.join(OUTDIR, f'pbf_{cfg}_{b}.docx')
        doc, ftr, styles_xml = build(b, cfg)
        with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', doc)
            z.writestr('word/styles.xml', styles_xml)
            z.writestr('word/footer1.xml', ftr)
    print('generated', len(cases), cfg, 'docs in', OUTDIR)


def measure(pat='pbf_*'):
    import glob
    import win32com.client, fitz
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        for p in sorted(glob.glob(os.path.join(OUTDIR, pat + '.docx'))):
            pdf = p[:-5] + '.pdf'
            if not os.path.exists(pdf):
                doc = word.Documents.Open(os.path.abspath(p), ReadOnly=True)
                doc.ExportAsFixedFormat(os.path.abspath(pdf), 17)
                doc.Close(False)
            d = fitz.open(pdf)
            bases = []
            for blk in d[0].get_text('dict')['blocks']:
                for l in blk.get('lines', []):
                    t = ''.join(s['text'] for s in l['spans'])
                    if t.strip().startswith('Item'):
                        bases.append(round(l['spans'][0]['origin'][1], 2))
            bases.sort()
            b = int(os.path.basename(p)[:-5].rsplit('_', 1)[-1])
            cfg = os.path.basename(p).split('_')[1]
            adv = 13.7988 if cfg in ('fN', 'fN2', 'fN3') else 12.6489
            print(f'{os.path.basename(p)[:-5]}: n={len(bases)} last={bases[-1]} '
                  f'(top72 + n*adv model: t_next={72 + len(bases) * adv:.2f})')
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        spec = sys.argv[2] if len(sys.argv) > 2 else 'coarse'
        if spec == 'coarse':
            for cfg in FOOTERS:
                gen(list(range(1200, 1601, 40)), cfg)
        elif spec == 'ny':
            for cfg in NY_CONFIGS:
                gen(list(range(1200, 1601, 40)), cfg)
        else:
            _, lo, hi, step, cfg = spec.split(':')
            gen(list(range(int(lo), int(hi) + 1, int(step))), cfg)
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pbf_*')
