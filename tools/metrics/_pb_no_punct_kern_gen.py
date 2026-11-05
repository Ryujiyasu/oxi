# -*- coding: utf-8 -*-
"""Does w:noPunctuationKerning stop the kern-driven pair halving?

ja/forms__03a9ef3d (blind-J p5-p6): settings doNotCompress + noPunctuationKerning,
Normal kern=1. Word sets «…という。）にわたる…» at a flat 10.68 per glyph;
Oxi halved «。» before «）» (S547: kern > 0 halves a closing pair even under
doNotCompress), fitted one more glyph per line and cascaded down p5-p6.

Faithful slice of that document (styles, theme, settings, sectPr): short
left-aligned paragraphs holding the pairs «。）» «、「» «」、» «）「», style a9
(ＭＳ 明朝 10.5, Normal kern 1). Arms: the document's noPunctuationKerning kept or
removed, and kern removed. Readout: each pair's first-glyph advance.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/forms/03a9ef3df7538948.docx'
OUT = 'tests/fixtures/no_punct_kern'; os.makedirs(OUT, exist_ok=True)
TEXTS = ['あいう。）えお', 'あいう、「えお', 'あいう」、えお', 'あいう）「えお']
ARMS = {'npk': (True, True), 'no_npk': (False, True), 'npk_nokern': (True, False)}


def build(name, keep_npk, kern):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                b0 = x.index('<w:body>') + len('<w:body>')
                sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>\s*</w:body>', x, re.S).group(0)
                k = '' if kern else '<w:kern w:val="0"/>'
                body = ''.join(f'<w:p><w:pPr><w:pStyle w:val="a9"/><w:jc w:val="left"/></w:pPr>'
                               f'<w:r><w:rPr>{k}</w:rPr><w:t>{t}</w:t></w:r></w:p>' for t in TEXTS)
                x = x[:b0] + body + sect + x[x.index('</w:body>') + len('</w:body>'):]
                data = x.encode('utf8')
            elif it.filename == 'word/settings.xml' and not keep_npk:
                data = data.decode('utf8').replace('<w:noPunctuationKerning/>', '').encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    paths = {n: build(n, *a) for n, a in ARMS.items()}
    if '--build' in sys.argv:
        sys.exit()
    import win32com.client, pymupdf
    for n, at in paths.items():
        app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(at[:-5] + '.pdf', 17)
            finally:
                d.Close(False)
        finally:
            app.Quit()
        page = pymupdf.open(at[:-5] + '.pdf')[0]
        lines = {}
        for t in page.get_texttrace():
            for c in t['chars']:
                lines.setdefault(round(c[2][1]), []).append((c[2][0], chr(c[0])))
        out = []
        for y in sorted(lines):
            cs = sorted(lines[y])
            s = ''.join(ch for _, ch in cs)
            if len(s) < 6:
                continue
            adv = [round(b[0] - a[0], 2) for a, b in zip(cs, cs[1:])]
            out.append(f'{s[3:5]} {adv[3]}/{adv[4]}')
        print(f'{n:11} ' + '  '.join(out), flush=True)
