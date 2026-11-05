# -*- coding: utf-8 -*-
"""Do a text box's paragraph after and the next paragraph's before add or collapse?

en/creative__045bc787 (blind-J): the side box's «Entrance Requirements»
(SidebarHeading, docDefaults after 200) followed by a bullet (before 200 from
docDefaults, after 0): Word's baseline gap 28.2, Oxi's 39.4 (= one 10pt more).

Faithful slice of that document (styles, theme, settings, sectPr, the side
text box): the box's content is replaced by «AAAA» (after a) + «BBBB» (before b)
+ «CCCC», Arial 12, line 240 auto. Arms (a, b) in twips. Readout: the baseline
gaps A->B and B->C from the Word PDF.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/en/creative/045bc787d68b33a5.docx'
OUT = 'tests/fixtures/txbx_collapse'; os.makedirs(OUT, exist_ok=True)
ARMS = {'a200_b200': (200, 200), 'a200_b0': (200, 0), 'a0_b200': (0, 200),
        'a400_b200': (400, 200), 'a200_b400': (200, 400), 'a0_b0': (0, 0)}


def para(text, before, after):
    return (f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}" w:line="240" w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>{text}</w:t></w:r></w:p>')


def build(name, a, b):
    zin = zipfile.ZipFile(SRC)
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                content = para('AAAA', 0, a) + para('BBBB', b, 0) + para('CCCC', 0, 0)
                n = 0

                def repl(m):
                    nonlocal n
                    if 'Entrance Requirements' in m.group(0):
                        n += 1
                        return '<w:txbxContent>' + content + '</w:txbxContent>'
                    return m.group(0)
                x = re.sub(r'<w:txbxContent>.*?</w:txbxContent>', repl, x, flags=re.S)
                assert n >= 1, n
                data = x.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    paths = {n: build(n, *ab) for n, ab in ARMS.items()}
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
        base = {}
        for t in page.get_texttrace():
            s = ''.join(chr(c[0]) for c in t['chars'])
            for k in ('AAAA', 'BBBB', 'CCCC'):
                if k in s:
                    base.setdefault(k, round(t['chars'][0][2][1], 2))
        ab = round(base['BBBB'] - base['AAAA'], 2) if 'AAAA' in base and 'BBBB' in base else None
        bc = round(base['CCCC'] - base['BBBB'], 2) if 'BBBB' in base and 'CCCC' in base else None
        print(f'{n:10} A {base.get("AAAA")}  A->B {ab}  B->C {bc}', flush=True)
