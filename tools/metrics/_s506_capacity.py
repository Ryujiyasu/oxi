# -*- coding: utf-8 -*-
"""S506 capacity diagnosis: a PURE-KANJI jc=left line (no punct → no hang/oidashi) at
compat 15, width ~450pt. Word fits 37 pure kanji (S492 §1). Measure Oxi's line-1 char
count vs Word — if Oxi < 37, the compat-15 capacity gap is a char-width/wrap-budget bug
(uniform); if Oxi == 37, the gap is punct/kinsoku-specific. cp932-safe."""
import os, sys, zipfile, subprocess, tempfile, io, json
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'oidashi')
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TEXT = '亜' * 60  # pure kanji, no punct; wraps several times

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
WRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings %s><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>' % (NS, compat))


RPR = '<w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝" w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'


def doc_xml():
    para = '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (RPR, TEXT)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1450" w:bottom="1440" w:left="1450" w:header="720" w:footer="720"/></w:sectPr>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, para, sect)


def build():
    p = os.path.join(OUT, 'cap_kanji_c15.docx')
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', WRELS); z.writestr('word/settings.xml', settings(15))
        z.writestr('word/document.xml', doc_xml())
    return p


def line1_count(glyphs, ykey):
    g = [x for x in glyphs if x['char'].strip()]
    rows = {}
    for x in g:
        rows.setdefault(round(x[ykey], 0), []).append(x)
    y0 = sorted(rows)[0]
    return len(rows[y0])


def main():
    dx = build()
    # Word
    wj = tempfile.mktemp(suffix='.json', dir='c:/tmp')
    subprocess.run([sys.executable, os.path.join(ROOT, 'tools', 'metrics', 'word_pdf_glyphs.py'), dx, wj], capture_output=True, timeout=120)
    W = json.load(io.open(wj, encoding='utf-8'))['pages'][0]['glyphs']
    # Oxi
    oj = tempfile.mktemp(suffix='.json', dir='c:/tmp')
    subprocess.run([DW, dx, tempfile.mktemp(dir='c:/tmp'), '150', '--dump-glyphs=' + oj], capture_output=True, timeout=120)
    O = json.load(io.open(oj, encoding='utf-8'))['pages'][0]['glyphs']
    wc = line1_count(W, 'y'); oc = line1_count(O, 'baseline')
    L = ['S506 capacity: pure-kanji jc=left compat15, width ~450pt (Word fits 37 per S492 §1)',
         'WORD line1 kanji = %d' % wc,
         'OXI  line1 kanji = %d' % oc,
         'VERDICT: %s' % ('Oxi capacity MATCHES Word (gap is punct/kinsoku-specific, not char-width)'
                          if oc == wc else 'Oxi capacity DIFFERS by %+d (char-width/wrap-budget bug)' % (oc - wc))]
    with io.open('c:/tmp/_s506_capacity_out.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print('\n'.join(L))


if __name__ == '__main__':
    main()
