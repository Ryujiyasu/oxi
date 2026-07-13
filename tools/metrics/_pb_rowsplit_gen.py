"""Word row-split FILL threshold — direct decision-rule probe.

uklocal p36/37: the row-2 split keeps 2 lines in Word / 3 in Oxi with
every reconstructible component (footer stack, table gap, row spans,
line heights, Segoe) confirmed EXACT to <=0.07pt — the decision differs
within one device quantum, unresolvable by position algebra (fitz ink
offsets wobble +-0.3). This probe measures the RULE directly:

  K filler lines (Arial 11, spacing 0) + exact spacer (line=X exact) +
  a 2-col bordered table (uklocal data-row replica: tcBorders sz6,
  tcMar top/bottom=105tw, cell paras direct before/after=120,
  line=0 atLeast, Arial 10) whose left cell has ONE paragraph of 6
  lines (ROWLINE1..6 via <w:br/>). No footer -> cbot = 769.9
  (pinned to 0.05 by _pb_fstack c0).

  X swept: readout = the highest ROWLINEn on page 1. Each n->n-1
  transition X pins Word's per-line fill threshold:
    line_n kept iff  table_top + lead + n*line + tail(?) <= cbot(+q?)
  with table_top = 72 + K*12.6489 + X/20 exactly. The transition PHASE
  gives the threshold formula to 0.1pt (full-line-box vs ink vs
  device-rounded), the transition SPACING confirms the per-line pitch.

Usage:
  python _pb_rowsplit_gen.py gen coarse | gen fine:LO:HI:STEP
  python _pb_rowsplit_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_rowsplit")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

R11 = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
R10 = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/>'
K = 48
ADV = 12.6489


def build(x):
    paras = []
    for i in range(K):
        paras.append(
            f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{R11}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R11}</w:rPr><w:t>Item {i:02d} alpha beta.</w:t></w:r></w:p>')
    paras.append(
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{x}" w:lineRule="exact"/>'
        f'<w:rPr>{R11}</w:rPr></w:pPr></w:p>')
    lines = '<w:r><w:br/></w:r>'.join(
        f'<w:r><w:rPr>{R10}</w:rPr><w:t>ROWLINE{n}</w:t></w:r>' for n in range(1, 7))
    cellp = (f'<w:p><w:pPr><w:spacing w:before="120" w:after="120" w:line="0" w:lineRule="atLeast"/>'
             f'<w:rPr>{R10}</w:rPr></w:pPr>{lines}</w:p>')
    cellq = (f'<w:p><w:pPr><w:spacing w:before="120" w:after="120" w:line="0" w:lineRule="atLeast"/>'
             f'<w:rPr>{R10}</w:rPr></w:pPr>'
             f'<w:r><w:rPr>{R10}</w:rPr><w:t>SideCell</w:t></w:r></w:p>')
    def tc(content):
        return ('<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>'
                '<w:tcBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                '<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                '<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tcBorders>'
                '<w:tcMar><w:top w:w="105" w:type="dxa"/><w:left w:w="105" w:type="dxa"/>'
                '<w:bottom w:w="105" w:type="dxa"/><w:right w:w="105" w:type="dxa"/></w:tcMar>'
                f'</w:tcPr>{content}</w:tc>')
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
           '<w:tblCellMar><w:top w:w="15" w:type="dxa"/><w:left w:w="15" w:type="dxa"/>'
           '<w:bottom w:w="15" w:type="dxa"/><w:right w:w="15" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>'
           f'<w:tr>{tc(cellp)}{tc(cellq)}</w:tr></w:tbl>')
    body = ''.join(paras) + tbl + '<w:p/>'
    body += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
             'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


def gen(cases):
    os.makedirs(OUTDIR, exist_ok=True)
    for x in cases:
        with zipfile.ZipFile(os.path.join(OUTDIR, f'prs_{x:04d}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/document.xml', build(x))
    print('generated', len(cases), 'docs in', OUTDIR)


def measure(pat='prs_*'):
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
            kept = []
            p1lines = {}
            for blk in d[0].get_text('dict')['blocks']:
                if blk.get('type') != 0:
                    continue
                for ln in blk['lines']:
                    t = ''.join(s['text'] for s in ln['spans'])
                    for n in range(1, 7):
                        if f'ROWLINE{n}' in t:
                            kept.append(n)
                            p1lines[n] = round(ln['bbox'][1], 2)
            x = int(os.path.basename(f)[:-5].rsplit('_', 1)[-1])
            top = 72 + K * ADV + x / 20.0
            last = max(kept) if kept else 0
            lasty = p1lines.get(last)
            print(f'prs_{x:04d}: kept={last} lastink={lasty} tbl_top_model={top:.2f}')
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        spec = sys.argv[2] if len(sys.argv) > 2 else 'coarse'
        if spec == 'coarse':
            gen(list(range(560, 941, 20)))
        else:
            _, lo, hi, step = spec.split(':')
            gen(list(range(int(lo), int(hi) + 1, int(step))))
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'prs_*')
