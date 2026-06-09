# -*- coding: utf-8 -*-
"""S518: does Word render the Symbol bullet (U+F0B7/U+2022) at the body baseline or offset?
Measure gen2_003 (and b5f706 '1)') bullet/marker baseline vs the next body glyph in Word PDF.
cp932-safe."""
import os, io, glob
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))

def measure(docx, marker_pred, label):
    import win32com.client, pythoncom, fitz
    pdf = os.path.join('c:/tmp', os.path.basename(docx)[:16] + '_w.pdf')
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(os.path.abspath(docx), ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    out = ['=== %s : %s' % (label, os.path.basename(docx))]
    doc = fitz.open(pdf); found = 0
    for pno in range(min(3, doc.page_count)):
        for blk in doc[pno].get_text('rawdict').get('blocks', []):
            for ln in blk.get('lines', []):
                chs = [c for sp in ln.get('spans', []) for c in sp.get('chars', [])]
                if not chs:
                    continue
                idxs = [i for i, c in enumerate(chs[:2]) if marker_pred(c['c'])]
                if not idxs:
                    continue
                mi = idxs[0]; mbl = chs[mi]['origin'][1]
                body_bl = None
                for c in chs[mi + 1:]:
                    if c['c'].strip():
                        body_bl = c['origin'][1]; bcp = hex(ord(c['c'])); break
                if body_bl is not None:
                    found += 1
                    if found <= 4:
                        out.append('  p%d marker=U+%04X bl=%.2f | body=%s bl=%.2f | dy=%+.2f' % (
                            pno + 1, ord(chs[mi]['c']), mbl, bcp, body_bl, mbl - body_bl))
    return out

def main():
    L = ['S518 Word marker/bullet baseline vs body']
    g2 = glob.glob(os.path.join(ROOT, 'tools/golden-test/documents/docx/gen2_003*.docx'))[0]
    L += measure(g2, lambda c: ord(c) in (0xF0B7, 0x2022, 0x25CF) or c == '⦁', 'gen2_003 bullet')
    b5 = glob.glob(os.path.join(ROOT, 'tools/golden-test/documents/docx/b5f706*.docx'))[0]
    L += measure(b5, lambda c: c == '1', 'b5f706 1)')
    txt = '\n'.join(L)
    io.open('c:/tmp/_s518_wordbullet.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)

if __name__ == '__main__':
    main()
