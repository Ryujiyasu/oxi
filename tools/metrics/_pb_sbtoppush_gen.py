"""Space-before at a page top -- PUSHED paragraphs vs natural overflow.

reference__0061531a p53 (2026-08-20 dissection): a 6-line subsection is sent
to the next page WHOLE by orphan control (1 line would have fit), and Word
renders the new page ~7.6pt lower than Oxi -- i.e. the space-before appears
to SURVIVE. But _pb_sbtop_gen (2026-07-13) pinned the plain natural-overflow
top as fully suppressed (y = margin). Hypothesis: Word distinguishes HOW the
paragraph reached the page top:

  overflow  : first line did not fit on the previous page -> suppress (known)
  orphan    : line 1 fit but widow/orphan control moved the whole paragraph
  keepNext  : the paragraph fit but was dragged by its follower
  keepLines : a split was possible but keepLines forced a whole-move

Arms (filler count n sweeps the boundary through every phase):
  nat  : 4-line target (3 internal <w:br/>), before swept 0/120/240/360/480.
         As n grows: fits -> 2+2 split -> orphan whole-push -> overflow.
  kn   : 1-line HEADING (keepNext, before=X) + 2-line follower.
  kl   : 4-line target with keepLines, before=X.
  pa   : nat grid but the LAST FILLER carries after=240 (excess-model test:
         does the push top apply max(0, before - prev_after)?).

Readout: page + ink y of TARGETLINE1 / HEADINGLINE (PDF via Word COM).
Phase is identified post hoc: which page line 1 landed on plus the y of the
before=0 control at the same n.

Usage:
  python _pb_sbtoppush_gen.py gen
  python _pb_sbtoppush_gen.py measure [pattern]
"""
import os, sys, zipfile

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_sbtoppush")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
            '</Relationships>')


def settings_xml(compat_mode):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings {W_NS}><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" '
            f'w:val="{compat_mode}"/></w:compat></w:settings>')

R = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
FILLER = 'Filler paragraph text line for the page fill sweep.'


def filler_para(i, after=0):
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="{after}" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{R}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R}</w:rPr><w:t>{FILLER} {i:02d}</w:t></w:r></w:p>')


def target_para(before, lines=4, keep_lines=False, keep_next=False, tag='TARGETLINE'):
    kl = '<w:keepLines/>' if keep_lines else ''
    kn = '<w:keepNext/>' if keep_next else ''
    runs = []
    for li in range(1, lines + 1):
        if li > 1:
            runs.append(f'<w:r><w:rPr>{R}</w:rPr><w:br/></w:r>')
        runs.append(f'<w:r><w:rPr>{R}</w:rPr><w:t>{tag}{li} probe text</w:t></w:r>')
    return (f'<w:p><w:pPr>{kn}{kl}'
            f'<w:spacing w:before="{before}" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{R}</w:rPr></w:pPr>{"".join(runs)}</w:p>')


def wrap_target_para(before):
    """A genuinely WRAPPING ~4-line paragraph (soft <w:br/> lines dodge
    widow/orphan control -- the splits observed in the first sweep prove it),
    with explicit widowControl so an orphan push can actually fire."""
    words = ' '.join(f'wrapfill{w:02d}' for w in range(40))
    text = f'TARGETLINE1 {words} probe end.'
    return (f'<w:p><w:pPr><w:widowControl/>'
            f'<w:spacing w:before="{before}" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>{R}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{R}</w:rPr><w:t>{text}</w:t></w:r></w:p>')


def build(arm, n_fill, before):
    paras = []
    last_after = 240 if arm == 'pa' else 0
    for i in range(n_fill):
        paras.append(filler_para(i, after=last_after if i == n_fill - 1 else 0))
    if arm == 'kn':
        paras.append(target_para(before, lines=1, keep_next=True, tag='HEADINGLINE'))
        paras.append(target_para(0, lines=2, tag='FOLLOWLINE'))
    elif arm == 'kl':
        paras.append(target_para(before, lines=4, keep_lines=True))
    elif arm.startswith('wrap'):
        paras.append(wrap_target_para(before))
    else:  # nat / pa / nat14
        paras.append(target_para(before, lines=4))
    body = ''.join(paras)
    body += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
             'w:left="1440" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


# Arial 11 single line ~12.65pt, A4 content ~697.6pt -> ~55 lines/page.
# nat grid covers split -> orphan-push -> overflow; controls at before=0.
ARM_COMPAT = {'wrap14': 14, 'wrap15': 15, 'nat14': 14}
CASES = (
    [('nat', n, b) for n in range(50, 57) for b in (0, 240)] +
    [('nat', n, b) for n in (52, 53, 54, 55) for b in (120, 360, 480)] +
    [('kn', n, b) for n in range(51, 56) for b in (0, 240)] +
    [('kl', n, b) for n in range(51, 56) for b in (0, 240)] +
    [('pa', n, 240) for n in (52, 53, 54, 55)]
)
# Round 2: a genuinely wrapping target (orphan control CAN fire) and
# compatibilityMode arms (the real doc that keeps its before is mode 14).
CASES_R2 = (
    [('wrap', n, b) for n in range(50, 57) for b in (0, 240)] +
    [('wrap14', n, b) for n in range(52, 56) for b in (0, 240)] +
    [('wrap15', n, b) for n in (53, 54) for b in (0, 240)] +
    [('nat14', n, b) for n in (54, 55) for b in (0, 240)]
)


def name(arm, n, b):
    return f'pbsp_{arm}_b{b:03d}_{n:02d}.docx'


def gen(cases=None):
    os.makedirs(OUTDIR, exist_ok=True)
    for arm, n, b in (cases or CASES):
        with zipfile.ZipFile(os.path.join(OUTDIR, name(arm, n, b)), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/document.xml', build(arm, n, b))
            cm = ARM_COMPAT.get(arm)
            if cm is not None:
                z.writestr('word/_rels/document.xml.rels', DOC_RELS)
                z.writestr('word/settings.xml', settings_xml(cm))
    print('generated', len(cases or CASES), 'docs in', OUTDIR)


def measure(pat='pbsp_*'):
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
            tgt = last_fill = None
            n_lines_p2 = 0
            for pi in range(len(d)):
                for blk in d[pi].get_text('dict')['blocks']:
                    if blk.get('type') != 0:
                        continue
                    for ln in blk['lines']:
                        t = ''.join(s['text'] for s in ln['spans'])
                        if ('TARGETLINE1' in t or 'HEADINGLINE1' in t) and tgt is None:
                            tgt = (pi + 1, round(ln['bbox'][1], 2))
                        if FILLER.split()[0] in t:
                            last_fill = (pi + 1, round(ln['bbox'][1], 2))
                        if pi == 1 and ('TARGETLINE' in t or 'FOLLOWLINE' in t
                                        or 'HEADINGLINE' in t):
                            n_lines_p2 += 1
            base = os.path.basename(f)[:-5]
            tp, ty = tgt if tgt else ('?', '?')
            lp, ly = last_fill if last_fill else ('?', '?')
            print(f'{base}: target p{tp} y={ty}  last_fill p{lp} y={ly}  p2_lines={n_lines_p2}')
    finally:
        word.Quit()


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        gen()
    elif mode == 'gen2':
        gen(CASES_R2)
    else:
        measure(sys.argv[2] if len(sys.argv) > 2 else 'pbsp_*')
