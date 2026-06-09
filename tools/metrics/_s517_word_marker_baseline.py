# -*- coding: utf-8 -*-
"""S517: confirm Word renders the circled-number list marker (U+2460..) at the SAME baseline as
the following body text on the same line (b837). Export full doc to PDF, find lines whose first
glyph is a circled number, report per-char baselines on that line. cp932-safe: UTF-8 file,
results to file, ASCII out (codepoints)."""
import os, io
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx',
                    'b837808d0555_20240705_resources_data_guideline_02.docx')
CIRCLED = set(range(0x2460, 0x2474))

def main():
    import win32com.client, pythoncom, fitz
    pdf = os.path.join('c:/tmp', 'b837_word.pdf')
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(os.path.abspath(DOCX), ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    doc = fitz.open(pdf)
    L = ['S517 Word b837 marker-vs-body baseline (lines whose text contains a circled number near start)']
    found = 0
    for pno in range(min(6, doc.page_count)):
        for blk in doc[pno].get_text('rawdict').get('blocks', []):
            for ln in blk.get('lines', []):
                chs = [c for sp in ln.get('spans', []) for c in sp.get('chars', [])]
                if not chs:
                    continue
                # circled char within first 3 glyphs
                idxs = [i for i, c in enumerate(chs[:3]) if ord(c['c']) in CIRCLED]
                if not idxs:
                    continue
                found += 1
                if found > 6:
                    break
                mi = idxs[0]
                marker_bl = chs[mi]['origin'][1]
                # next non-space body glyph
                body_bl = None
                for c in chs[mi + 1:]:
                    if c['c'].strip():
                        body_bl = c['origin'][1]; body_cp = hex(ord(c['c'])); break
                L.append('p%d marker=U+%04X bl=%.2f | next_body=%s bl=%s | dy(marker-body)=%s | x_marker=%.1f' % (
                    pno + 1, ord(chs[mi]['c']), marker_bl,
                    body_cp if body_bl else '-', ('%.2f' % body_bl) if body_bl else '-',
                    ('%+.2f' % (marker_bl - body_bl)) if body_bl else '-', chs[mi]['origin'][0]))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s517_word_marker.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)

if __name__ == '__main__':
    main()
