"""pgNumType w:start — page-number restart derivation.

Oxi's PAGE field post-pass (layout/mod.rs) substitutes `page_idx + 1`, i.e.
the PHYSICAL page index, and never reads `Page.page_number_start` (parsed
since the section-properties work but with no consumer). Word instead
numbers each page within its SECTION's numbering sequence:
a section declaring `<w:pgNumType w:start="N"/>` restarts at N.

Render-truth already confirms the bug on albaluna2col_3227 (start=3,
1 page): Word's footer shows "3", Oxi shows "1".

This sweep pins the rest of the rule with controls. Each doc has two
2-page sections; every page carries a footer "PG=<PAGE field>":

  base       : no pgNumType anywhere            -> control, expect 1,2,3,4
  s1start5   : sec1 start=5, sec2 none          -> continue? 5,6,7,8 vs 5,6,1,2
  s2start1   : sec1 none, sec2 start=1          -> restart:  1,2,1,2
  s2start10  : sec1 none, sec2 start=10         -> restart:  1,2,10,11
  s2fmtonly  : sec2 pgNumType fmt, NO start     -> does fmt alone reset?
  cont_s2s7  : sec2 CONTINUOUS, start=7         -> merged-page (S560) case

Readout: per PDF page, the footer's PG= number and the body marker.

Usage:
  python _pb_pgnum_gen.py gen
  python _pb_pgnum_gen.py measure [pattern]
"""
import os
import sys
import zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_pgnum")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
            '</Relationships>')

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'

# Footer carrying a live PAGE field, prefixed so the PDF readout is unambiguous.
# NOTE the trailing space in "PG= ": without it the prefix run and the field run
# are adjacent, and Oxi's word accumulator merges them into ONE fragment keeping
# the FIRST run's style — the field_type is lost and the PAGE substitution never
# fires (Oxi renders a literal "PG=#"). That is a real, SEPARATE bug in the
# S826/S899 fragment-flattening family; the space keeps it out of this probe.
FOOTER = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:ftr {W_NS}><w:p><w:pPr><w:rPr>{R}</w:rPr></w:pPr>'
          f'<w:r><w:rPr>{R}</w:rPr><w:t xml:space="preserve">PG= </w:t></w:r>'
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
          '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
          f'<w:r><w:rPr>{R}</w:rPr><w:t>9</w:t></w:r>'
          '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
          '</w:p></w:ftr>')

PGSZ = ('<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
        'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/>')
FTR_REF = '<w:footerReference w:type="default" r:id="rId10"/>'


def para(text, page_break=False):
    brk = '<w:r><w:br w:type="page"/></w:r>' if page_break else ''
    return (f'<w:p><w:pPr><w:rPr>{R}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R}</w:rPr><w:t>{text}</w:t></w:r>{brk}</w:p>')


def sectpr(pgnum, cont=False):
    """Section properties: footer ref + optional pgNumType + optional continuous."""
    typ = '<w:type w:val="continuous"/>' if cont else '<w:type w:val="nextPage"/>'
    return f'<w:sectPr>{FTR_REF}{typ}{PGSZ}{pgnum}</w:sectPr>'


def build(sec1_pgnum, sec2_pgnum, sec2_cont=False):
    # Section 1: two pages (marker A, page break, marker B), sectPr in a pPr.
    body = para('SEC1PAGEA', page_break=True)
    body += para('SEC1PAGEB')
    body += f'<w:p><w:pPr>{sectpr(sec1_pgnum)}</w:pPr></w:p>'
    # Section 2: two pages; its sectPr is the body-level one.
    body += para('SEC2PAGEA', page_break=True)
    body += para('SEC2PAGEB')
    body += sectpr(sec2_pgnum, cont=sec2_cont)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS} {R_NS}><w:body>{body}</w:body></w:document>')


PGNUM_START = '<w:pgNumType w:start="{}"/>'

CASES = {
    'base':      dict(sec1_pgnum='', sec2_pgnum=''),
    's1start5':  dict(sec1_pgnum=PGNUM_START.format(5), sec2_pgnum=''),
    's2start1':  dict(sec1_pgnum='', sec2_pgnum=PGNUM_START.format(1)),
    's2start10': dict(sec1_pgnum='', sec2_pgnum=PGNUM_START.format(10)),
    's2fmtonly': dict(sec1_pgnum='', sec2_pgnum='<w:pgNumType w:fmt="lowerRoman"/>'),
    'cont_s2s7': dict(sec1_pgnum='', sec2_pgnum=PGNUM_START.format(7), sec2_cont=True),
}


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for name, kw in CASES.items():
        path = os.path.join(OUTDIR, f'pgn_{name}.docx')
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOC_RELS)
            z.writestr('word/footer1.xml', FOOTER)
            z.writestr('word/document.xml', build(**kw))
    print('generated', len(CASES), 'docs in', os.path.abspath(OUTDIR))


def measure(pat='pgn_*'):
    import glob
    import re
    import win32com.client
    import fitz
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
            out = []
            for pi in range(len(d)):
                txt = d[pi].get_text()
                pg = re.search(r'PG=\s*(\S+)', txt)
                mk = re.search(r'(SEC\dPAGE[AB])', txt)
                out.append(f'{mk.group(1) if mk else "?":10s}->{pg.group(1) if pg else "(none)"}')
            print(f'{os.path.basename(f)[:-5]:14s} pages={len(d)}  ' + '  '.join(out))
            d.close()
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        gen()
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pgn_*')
