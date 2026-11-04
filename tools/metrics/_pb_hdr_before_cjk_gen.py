# -*- coding: utf-8 -*-
"""Does a CJK document's header paragraph space-before push the body down?

blind-J ja/reports__75ebb654: the default header is one empty paragraph with
before=120 (6pt); header distance 48.2, top margin 56.7. Word's body sits 6.15pt
lower than Oxi's once the header line is measured in the substituted face --
S813 (header before/after) is gated to documents without real CJK.

Faithful slice of that document (its header part, styles, settings): the
header paragraph's before = 0 / 6 / 12 / 24pt. The header style carries
beforeLines=50, which outranks a bare w:before, so each arm writes
beforeLines=0 with it (the first run wrote w:before alone and every arm read
99.38). Readout: the body's first baseline («2025-01-13»).

Result (2026-11-04): baseline = 93.38 + before (0 / 6 / 12 / 24 -> 93.38 /
99.38 / 105.38 / 117.38); no header part -> Word makes one in the header style
(beforeLines 50 = 6pt) -> 99.38. The header stack carries its first
paragraph's before in a CJK document too (OXI_S813_CJK).
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/reports/75ebb6540d49b937.docx'
OUT = 'tests/fixtures/hdr_before_cjk'; os.makedirs(OUT, exist_ok=True)
ARMS = [0, 120, 240, 480, 'nohdr']
import sys as _s
if '--only-nohdr' in _s.argv: ARMS = ['nohdr']


def build(before):
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    at = os.path.abspath(os.path.join(OUT, f'b{before}.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, data in items:
            if before == 'nohdr' and it.filename == 'word/document.xml':
                data = re.sub(r'<w:headerReference [^>]*/>', '', data.decode('utf8')).encode('utf8')
            if re.match(r'word/header\d*\.xml', it.filename) and before != 'nohdr':
                h = data.decode('utf8')
                h = h.replace('<w:spacing w:before="120"/>', f'<w:spacing w:beforeLines="0" w:before="{before}"/>')
                data = h.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    for b in ARMS:
        at = build(b)
        app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
            finally:
                d.Close(False)
        finally:
            app.Quit()
        P = pymupdf.open(at[:-5] + '.pdf')[0]
        y = next(s['origin'][1] for blk in P.get_text('dict')['blocks'] for l in blk.get('lines', [])
                 for s in l['spans'] if '2025-01-13' in s['text'])
        print(f'before {b}: body first baseline {y:.2f}', flush=True)
