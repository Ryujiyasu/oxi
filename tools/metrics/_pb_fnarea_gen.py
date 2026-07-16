# -*- coding: utf-8 -*-
"""Footnote AREA placement/roll-over + built-in separator stack probe.

81e80 anatomy (Word render truth): the body limit obeys the FULL-reservation
model (limit = margin_bottom_line - sep - SUM(all committed note heights):
wi=23 pushes at 570.3 > 557.7 and L9 keeps at 556.5 <= 557.7 with sep ~= 5.2),
while the PLACEMENT bottom-packs whole notes in ref order and ROLLS the tail
(16/17/18) to the NEXT page's area head. The open constants:
  (a) the built-in separator RESERVATION stack for the no-declared-separator,
      small-fn class (81e80 implies ~5.2pt; uklocal's DECLARED sep = 13.43;
      bunkacontract's no-grid = one fn line ~ its fn line height)
  (b) the placement cutoff (which notes fit whole above the margin line).

Shape: TNR 12 body (Normal, no spacing), NO docGrid, Letter, 1440 margins.
  K filler lines; then 14 one-ref paras (fills the area like 81e80's 2..15);
  then the TARGET para whose LAST line carries 4 refs; then one plain para.
  Footnotes: single short 8pt TNR lines (styled like 81e80's).
  An exact-line spacer (line=X lineRule=exact) above the target sweeps the
  target's bottom position in 2tw steps across the keep/push flip.

Read (Word PDF via fitz): per-page body last line, note ids per page, first
note y, rule y. The X where the target's last line flips pages pins the
reservation stack to 0.1pt; the note-id split per page pins the placement.

Usage: gen [--sweep lo hi step] | measure
"""
import os, sys, zipfile, json, glob

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_fnarea")

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
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
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="FnText"><w:name w:val="footnote text"/>'
          '<w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>'
          '<w:style w:type="character" w:styleId="FnRef"><w:name w:val="footnote reference"/>'
          '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
          '</w:styles>')

FILLER = ('Lorem ipsum dolor sit amet consectetur adipiscing elit sed do '
          'eiusmod tempor incididunt ut labore et dolore.')


def ref_run(fid):
    return (f'<w:r><w:rPr><w:rStyle w:val="FnRef"/></w:rPr>'
            f'<w:footnoteReference w:id="{fid}"/></w:r>')


def para(text, tail_refs=(), spacing=''):
    runs = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
    for fid in tail_refs:
        runs += ref_run(fid)
        runs += '<w:r><w:t xml:space="preserve">; and more</w:t></w:r>'
    return f'<w:p><w:pPr>{spacing}</w:pPr>{runs}</w:p>'


def footnotes(n):
    fns = ('<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
           '<w:r><w:separator/></w:r></w:p></w:footnote>'
           '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
           '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>')
    for i in range(1, n + 1):
        fns += (f'<w:footnote w:id="{i}"><w:p><w:pPr><w:pStyle w:val="FnText"/></w:pPr>'
                f'<w:r><w:rPr><w:rStyle w:val="FnRef"/></w:rPr><w:footnoteRef/></w:r>'
                f'<w:r><w:t xml:space="preserve"> NOTE{i} ref text.</w:t></w:r></w:p></w:footnote>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:footnotes {W_NS}>{fns}</w:footnotes>')


def build(name, spacer_tw):
    body = []
    nfill = int(os.environ.get('FNA_FILL', '17'))
    for k in range(nfill):
        body.append(para(f'F{k:02d} ' + FILLER))
    # exact spacer BEFORE the ref block so the sweep slides the whole
    # R-block (and the page-boundary flip line) in 0.1pt steps.
    body.append(para('SPACER', spacing=f'<w:spacing w:line="{spacer_tw}" w:lineRule="exact"/>'))
    # 14 one-ref paragraphs (notes 1..14)
    for i in range(1, 15):
        body.append(para(f'R{i:02d} short reference line with a citation', (i,)))
    # target: last line carries refs 15..18
    body.append(para('TARGET begins here with a long lead so the final line '
                     'carries the four citations Bates', (15, 16, 17, 18)))
    body.append(para('AFTER paragraph plain text.'))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {W_NS}><w:body>{"".join(body)}'
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(os.path.join(OUTDIR, name), 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOCRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/footnotes.xml', footnotes(18))


def gen(sweep):
    os.makedirs(OUTDIR, exist_ok=True)
    for x in sweep:
        build(f'fna_{x:05d}.docx', x)
    print('generated', len(sweep), 'docs in', os.path.abspath(OUTDIR))


def measure():
    import win32com.client
    import fitz
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for f in sorted(glob.glob(os.path.join(OUTDIR, 'fna_*.docx'))):
            pdf = f[:-5] + '.pdf'
            if not os.path.exists(pdf):
                d = word.Documents.Open(os.path.abspath(f), ReadOnly=True)
                d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
                d.Close(False)
            doc = fitz.open(pdf)
            info = []
            for pi, pg in enumerate(doc):
                lines = []
                for blk in pg.get_text('dict')['blocks']:
                    for ln in blk.get('lines', []):
                        t = ''.join(s['text'] for s in ln['spans']).strip()
                        if t:
                            lines.append((round(ln['bbox'][1], 1), t))
                lines.sort()
                notes = [t.split()[1] for y, t in lines if 'NOTE' in t and 'ref text' in t]
                body = [(y, t[:30]) for y, t in lines if 'NOTE' not in t]
                rules = [round(dr['rect'].y0, 1) for dr in pg.get_drawings()
                         if dr['rect'].width > 50 and dr['rect'].height < 3 and dr['rect'].y0 > 300]
                first_note_y = min((y for y, t in lines if 'NOTE' in t), default=None)
                info.append({'page': pi + 1, 'notes': notes, 'rule': rules,
                             'first_note_y': first_note_y,
                             'last_body': body[-1] if body else None,
                             'target_page': next((pi + 1 for y, t in lines if t.startswith('TARGET')), None)})
            doc.close()
            base = os.path.basename(f)[:-5]
            # summary: which page holds TARGET's tail line + per-page note split
            tgt = [p['page'] for p in info if any('TARGET' in (p['last_body'] or ('', ''))[1] for _ in [0])]
            res[base] = info
            per = {p['page']: p['notes'] for p in info if p['notes']}
            print(f"{base}: notes/page={per}")
            for p in info:
                if p['first_note_y']:
                    print(f"   p{p['page']} rule={p['rule']} note1_y={p['first_note_y']} last_body={p['last_body']}")
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, '_measure.json'), 'w'), indent=1)


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        if '--sweep' in sys.argv:
            i = sys.argv.index('--sweep')
            lo, hi, st = int(sys.argv[i+1]), int(sys.argv[i+2]), int(sys.argv[i+3])
            sweep = list(range(lo, hi + 1, st))
        else:
            sweep = [240, 300, 360, 420, 480, 540, 600]
        gen(sweep)
    else:
        measure()
