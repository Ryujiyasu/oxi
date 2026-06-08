# -*- coding: utf-8 -*-
"""S502 cellpos measurement: for each cellpos variant, compare Word's first-char x
(COM->PDF->fitz) vs Oxi's first-char x (dwrite --dump-glyphs). The col1 content is
'あいう' (short, NON-wrapping) so line_total_w is identical across variants -> any x
delta is purely the firstLine / jc=center interaction. cp932-safe: ASCII output to
a file, no Japanese console eyeballing. NEEDLE = first CJK glyph after the REF para."""
import os, sys, json, subprocess, tempfile
import fitz
import win32com.client, pythoncom

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
REPRO = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'cellpos')
NEEDLE = 'あ'  # first char of the col1 content para
DPI = 150


def word_first_x(docx):
    docx = os.path.abspath(docx)
    pdf = os.path.splitext(docx)[0] + '_rt.pdf'
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Open(docx, ReadOnly=True)
        doc.ExportAsFixedFormat(pdf, 17)
        doc.Close(False)
    finally:
        word.Quit()
    d = fitz.open(pdf)
    page = d[0]
    rd = page.get_text('rawdict')
    for blk in rd.get('blocks', []):
        for line in blk.get('lines', []):
            for span in line.get('spans', []):
                for ch in span.get('chars', []):
                    if ch['c'] == NEEDLE:
                        return round(ch['origin'][0], 2)
    return None


def oxi_first_x(docx):
    fd, jp = tempfile.mkstemp(suffix='.json', dir='c:/tmp'); os.close(fd)
    subprocess.run([DW, os.path.abspath(docx), tempfile.mktemp(dir='c:/tmp'), str(DPI),
                    '--dump-glyphs=' + jp], capture_output=True, timeout=300)
    d = json.load(open(jp, encoding='utf-8')); os.unlink(jp)
    for page in d['pages']:
        for g in page['glyphs']:
            if g['char'] == NEEDLE:
                return round(g['x'], 2)
    return None


def main():
    variants = ['cp_center_fl.docx', 'cp_left_fl.docx', 'cp_center_nofl.docx', 'cp_left_nofl.docx']
    out = 'c:/tmp/_s502_cellpos_out.txt'
    lines = ['S502 cellpos: Word vs Oxi first-char (%s) x; content=AIU (no wrap)' % NEEDLE,
             '%-22s %10s %10s %8s' % ('variant', 'Word_x', 'Oxi_x', 'dx')]
    for v in variants:
        p = os.path.join(REPRO, v)
        if not os.path.exists(p):
            lines.append('%-22s MISSING' % v); continue
        wx = word_first_x(p)
        ox = oxi_first_x(p)
        dx = (ox - wx) if (wx is not None and ox is not None) else None
        lines.append('%-22s %10s %10s %8s' % (
            v, ('%.2f' % wx if wx is not None else 'NA'),
            ('%.2f' % ox if ox is not None else 'NA'),
            ('%+.2f' % dx if dx is not None else 'NA')))
    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print('wrote', out)


if __name__ == '__main__':
    main()
