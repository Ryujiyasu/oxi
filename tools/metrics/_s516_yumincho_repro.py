# -*- coding: utf-8 -*-
"""S516 faithful db9ca top-structure repro using the REAL theme font (Yu Mincho / 游明朝)
at sz21 body, to reproduce the +3.3pt title->Note offset that the Times-New-Roman repro
(S513) missed. Variants decompose where the offset enters (empty-p2 vs body line vs title).
Builds docx, exports via Word -> PDF baselines, runs Oxi dwrite dump-glyphs at dpi72,
reports per-line Oxi-Word. cp932-safe: UTF-8 file, results to file, ASCII verdicts."""
import os, zipfile, subprocess, io, json
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'emptypara')
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
os.makedirs(OUT, exist_ok=True)
YU = '游明朝'  # Yu Mincho
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')

def rpr(sz):
    return ('<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/><w:sz w:val="%d"/></w:rPr>'
            % (YU, YU, YU, sz))

def para(sz, text, exline=0):
    ppr = ('<w:pPr><w:spacing w:line="%d" w:lineRule="exact"/></w:pPr>' % exline) if exline else ''
    if text == '':
        return '<w:p><w:pPr>%s</w:pPr></w:p>' % rpr(sz)
    return '<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ppr, rpr(sz), text)

SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397"/>'
        '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')

# Body lines: short ASCII so each is one line; T/N/H prefixes are unique anchors.
def doc(paras):
    body = ''.join(paras)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s</w:body></w:document>'
            % (NS, body, SECT))

VARIANTS = {
    # faithful: empty24 / title28 exact420 / empty24 / 3 body lines (sz21)
    'yu_full.docx': [para(24, ''), para(28, 'TITLEXX', 420), para(24, ''),
                     para(21, 'Note line one'), para(21, 'Body line two'), para(21, 'Body line three')],
    # no empty-p2 (title directly to body): isolates empty-p2 contribution
    'yu_noempty2.docx': [para(24, ''), para(28, 'TITLEXX', 420),
                         para(21, 'Note line one'), para(21, 'Body line two'), para(21, 'Body line three')],
    # no exact title (normal title): isolates the exact-line phase
    'yu_noexact.docx': [para(24, ''), para(28, 'TITLEXX'), para(24, ''),
                        para(21, 'Note line one'), para(21, 'Body line two'), para(21, 'Body line three')],
    # body only (no title region): pure body grid baseline
    'yu_bodyonly.docx': [para(21, 'Note line one'), para(21, 'Body line two'), para(21, 'Body line three')],
}

def build(name, paras):
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc(paras))
    return p

def word_lines(dx):
    import win32com.client, pythoncom, fitz
    dx = os.path.abspath(dx); pdf = os.path.splitext(dx)[0] + '_rt.pdf'
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(dx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    lines = []
    for blk in fitz.open(pdf)[0].get_text('rawdict').get('blocks', []):
        for ln in blk.get('lines', []):
            chs = [c for sp in ln.get('spans', []) for c in sp.get('chars', [])]
            txt = ''.join(c['c'] for c in chs).strip()
            if txt:
                lines.append((txt, chs[0]['origin'][1]))
    return lines

def oxi_lines(dx):
    pre = os.path.join('c:/tmp', os.path.splitext(os.path.basename(dx))[0] + '_oxi')
    gj = pre + '_glyphs.json'
    subprocess.run([EXE, os.path.abspath(dx), pre, '72', '--dump-glyphs=' + gj], capture_output=True, text=True)
    g = sorted(json.load(open(gj, encoding='utf-8'))['pages'][0]['glyphs'], key=lambda c: (c['baseline'], c['x']))
    lines = []; cur = []
    for x in g:
        if cur and abs(x['baseline'] - cur[0]['baseline']) > 8:
            lines.append(cur); cur = []
        cur.append(x)
    if cur: lines.append(cur)
    res = []
    for ln in lines:
        import collections
        bl = collections.Counter(round(c['baseline'] * 2) / 2 for c in ln).most_common(1)[0][0]
        res.append((''.join(c['char'] for c in sorted(ln, key=lambda c: c['x'])).strip(), bl))
    return res

def main():
    L = ['S516 Yu Mincho faithful db9ca-top repro (Oxi dwrite dpi72 - Word PDF), px==pt']
    for name, paras in VARIANTS.items():
        dx = build(name, paras)
        wl = word_lines(dx); ol = oxi_lines(dx)
        L.append('')
        L.append('=== %s  (n_word=%d n_oxi=%d)' % (name, len(wl), len(ol)))
        n = min(len(wl), len(ol))
        for i in range(n):
            wt, wy = wl[i]; ot, oy = ol[i]
            flag = '' if wt[:3] == ot[:3] else ' MISMATCH(%r|%r)' % (wt[:6], ot[:6])
            L.append('  %d w=%7.2f o=%7.2f  d=%+6.2f%s' % (i, wy, oy, oy - wy, flag))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s516_out.txt', 'w', encoding='utf-8').write(txt + '\n')
    # ASCII-only console
    for line in txt.split('\n'):
        try: print(line)
        except Exception: print(line.encode('ascii', 'replace').decode())

if __name__ == '__main__':
    main()
