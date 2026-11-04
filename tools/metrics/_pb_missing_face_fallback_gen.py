# -*- coding: utf-8 -*-
"""Which face does Word draw for a run whose font is not installed, when the
font table gives it nothing to go on (PANOSE all zero, charset 0)?

blind-J en/correspondence__00f9eabae: body runs in «Twinkl Cursive Looped»
(fontTable: panose 02000000000000000000, charset 00, family auto, sig usb0 3)
-- not installed. Word draws Calibri (the theme minor face; docDefaults say
Times New Roman); Oxi keeps the unknown family, measures with a fallback and
the renderer draws something else.

Faithful slice of that document (styles, theme, settings, fontTable): one
paragraph per arm, the run's ascii/hAnsi face = an unknown name whose
fontTable entry is Twinkl's / absent / a serif PANOSE (Times-like) / a
sans PANOSE (Arial-like). Readout: the PDF span font of each arm.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/en/correspondence/00f9eabae239259f.docx'
OUT = 'tests/fixtures/missing_face_fallback'; os.makedirs(OUT, exist_ok=True)
ARMS = {
    'twinkl_entry': ('Zqx Twinkl Like', '02000000000000000000', 'auto', 'variable'),
    'no_entry': ('Zqx No Entry', None, None, None),
    'serif_panose': ('Zqx Serif Like', '02020603050405020304', 'roman', 'variable'),
    'sans_panose': ('Zqx Sans Like', '020B0604020202020204', 'swiss', 'variable'),
    'auto_nopan_var': ('Zqx Auto NoPan Var', '', 'auto', 'variable'),
    'auto_pan0_default': ('Zqx Auto Pan0 Def', '02000000000000000000', 'auto', 'default'),
    'auto_nopan_default': ('Zqx Auto NoPan Def', '', 'auto', 'default'),
    'auto_s1444_usb1_28': ('Zqx Auto Usb28', '', 'auto', 'default', '10000000'),
    'twinkl_usb1_28': ('Zqx Twinkl Usb28', '02000000000000000000', 'auto', 'variable', '10000000'),
}


def build():
    zin = zipfile.ZipFile(SRC)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    at = os.path.abspath(os.path.join(OUT, 'arms.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, data in items:
            if it.filename == 'word/document.xml':
                x = data.decode('utf8')
                b0 = x.index('<w:body>') + len('<w:body>')
                sect = re.search(r'<w:sectPr[ >].*?</w:sectPr>\s*</w:body>', x, re.S).group(0)
                body = ''.join(
                    f'<w:p><w:r><w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}"/><w:sz w:val="24"/></w:rPr>'
                    f'<w:t>{k} The quick brown fox jumps</w:t></w:r></w:p>'
                    for k, (face, *_r) in ARMS.items())
                x = x[:b0] + body + sect + x[x.index('</w:body>') + len('</w:body>'):]
                data = x.encode('utf8')
            if it.filename == 'word/fontTable.xml':
                f = data.decode('utf8')
                add = ''
                for face, pan, fam, pitch, *u in ARMS.values():
                    usb1 = u[0] if u else '00000001'
                    if fam:
                        add += (f'<w:font w:name="{face}">' + (f'<w:panose1 w:val="{pan}"/>' if pan else '') + '<w:charset w:val="00"/>'
                                f'<w:family w:val="{fam}"/><w:pitch w:val="{pitch}"/>'
                                f'<w:sig w:usb0="00000003" w:usb1="{usb1}" w:usb2="00000000" w:usb3="00000000" '
                                'w:csb0="00000001" w:csb1="00000000"/></w:font>')
                f = f.replace('</w:fonts>', add + '</w:fonts>')
                data = f.encode('utf8')
            z.writestr(it, data)
    return at


if __name__ == '__main__':
    import win32com.client, pymupdf
    at = build()
    if '--build' in sys.argv:
        sys.exit()
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
    for b in P.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            s = l['spans'][0]
            if 'quick' in ''.join(sp['text'] for sp in l['spans']):
                end = l['spans'][-1]['bbox'][2]
                print(f'{s["text"][:14]:16} font {s["font"]:22} width {end - s["bbox"][0]:.2f}')
