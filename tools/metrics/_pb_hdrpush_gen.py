"""Header pushdown + STYLEREF-in-header resolution — derivation probes.

reference__0061531a p52 (2026-08-20): Word re-evaluates the header's STYLEREF
fields per page; the resolved Part name wraps to 2 lines, the header bottom
crosses the top margin, and the BODY starts ~7pt below the margin. Oxi draws
the cached field result (1 line shorter) -> no pushdown -> one orphan decision
flips -> the doc's single Phase-1 slip. Two questions must be pinned before
implementing:

  push arm : body_top = f(header height)? Sweep K fixed header lines across
             the margin boundary (+ an after-spacing arm) with NO fields, so
             Oxi can be run on the same docs (is Oxi's s755 pushdown exact?).
  ref arm  : which paragraph does a header STYLEREF resolve to?
             H1 'AAA' on p1, H1 'BBB' mid p3: p1 -> AAA (first on page?),
             p2 -> AAA (last before?), p3 -> BBB (first on page, mid-page).
             refN variant: first H1 occurs on p2 (nothing before p1).

Usage:
  python _pb_hdrpush_gen.py gen
  python _pb_hdrpush_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_hdrpush")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
            '</Relationships>')

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          # custom name: a Japanese Word localizes built-in names ("heading 1"
          # -> STYLEREF error), custom names are looked up verbatim
          '<w:style w:type="paragraph" w:customStyle="1" w:styleId="ProbeHead">'
          '<w:name w:val="ProbeHead"/>'
          '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="22"/></w:rPr>'
          '</w:style></w:styles>')

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
FILLER = 'Filler paragraph text line for the page fill sweep.'


def filler(i):
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{R}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R}</w:rPr><w:t>{FILLER} {i:02d}</w:t></w:r></w:p>')


def header_fixed(k_lines, last_after=0):
    ps = []
    for i in range(k_lines):
        after = last_after if i == k_lines - 1 else 0
        ps.append(f'<w:p><w:pPr><w:spacing w:before="0" w:after="{after}" w:line="240" w:lineRule="auto"/>'
                  f'<w:rPr>{R}</w:rPr></w:pPr>'
                  f'<w:r><w:rPr>{R}</w:rPr><w:t>HDR line {i:02d}</w:t></w:r></w:p>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:hdr {W_NS}>{"".join(ps)}</w:hdr>')


HEADER_REF = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              f'<w:hdr {W_NS}>'
              '<w:p><w:pPr><w:rPr>' + R + '</w:rPr></w:pPr>'
              '<w:r><w:rPr>' + R + '</w:rPr><w:t xml:space="preserve">REF[</w:t></w:r>'
              '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
              '<w:r><w:instrText xml:space="preserve"> STYLEREF "ProbeHead" </w:instrText></w:r>'
              '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
              '<w:r><w:rPr>' + R + '</w:rPr><w:t>CACHEDXX</w:t></w:r>'
              '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
              '<w:r><w:rPr>' + R + '</w:rPr><w:t>]END</w:t></w:r>'
              '</w:p></w:hdr>')


def h1(text):
    return (f'<w:p><w:pPr><w:pStyle w:val="ProbeHead"/></w:pPr>'
            f'<w:r><w:t>{text}</w:t></w:r></w:p>')


def build(arm, k, after):
    if arm == 'push':
        body = ''.join(filler(i) for i in range(3))
        hdr = header_fixed(k, after)
    elif arm == 'ref':
        # H1 AAA at doc start (p1) / H1 BBB mid page 3 (after 120 fillers).
        body = (h1('AAAFIRST heading') +
                ''.join(filler(i) for i in range(120)) +
                h1('BBBSECOND heading') +
                ''.join(filler(200 + i) for i in range(10)))
        hdr = HEADER_REF
    else:  # refN: first H1 appears on page 2 only
        body = (''.join(filler(i) for i in range(60)) +
                h1('CCCLATE heading') +
                ''.join(filler(200 + i) for i in range(10)))
        hdr = HEADER_REF
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS} {R_NS}><w:body>{body}'
           '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/>'
           '<w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
           'w:left="1440" w:header="720" w:footer="709" w:gutter="0"/></w:sectPr>'
           '</w:body></w:document>')
    return doc, hdr


CASES = ([('push', k, 0) for k in range(1, 8)] +
         [('push', 3, 240), ('push', 5, 240)] +
         [('ref', 0, 0), ('refN', 0, 0)])


def name(arm, k, a):
    return f'pbh_{arm}_{k}_{a:03d}.docx'


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for arm, k, a in CASES:
        doc, hdr = build(arm, k, a)
        with zipfile.ZipFile(os.path.join(OUTDIR, name(arm, k, a)), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOC_RELS)
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/document.xml', doc)
            z.writestr('word/header1.xml', hdr)
    print('generated', len(CASES), 'docs in', OUTDIR)


def measure(pat='pbh_*'):
    import glob
    import win32com.client, fitz
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        for f in sorted(glob.glob(os.path.join(OUTDIR, pat + '.docx'))):
            pdf = f[:-5] + '.pdf'
            if not os.path.exists(pdf):
                doc = word.Documents.Open(os.path.abspath(f), ReadOnly=True)
                doc.ExportAsFixedFormat(os.path.abspath(pdf), 17)
                doc.Close(False)
            d = fitz.open(pdf)
            base = os.path.basename(f)[:-5]
            if '_push_' in base:
                got = []
                for blk in d[0].get_text('dict')['blocks']:
                    if blk.get('type') != 0:
                        continue
                    for ln in blk['lines']:
                        t = ''.join(s['text'] for s in ln['spans'])
                        y = round(ln['bbox'][1], 2)
                        if 'Filler' in t and ' 00' in t:
                            got.append(('body0', y))
                        if 'HDR line' in t:
                            got.append((t.strip()[-2:], y))
                print(f'{base}: {got}')
            else:
                for pi in range(min(len(d), 4)):
                    for blk in d[pi].get_text('dict')['blocks']:
                        if blk.get('type') != 0:
                            continue
                        for ln in blk['lines']:
                            t = ''.join(s['text'] for s in ln['spans'])
                            if 'REF[' in t:
                                print(f'{base}: p{pi+1} header = {t.strip()[:60]!r}')
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        gen()
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pbh_*')
