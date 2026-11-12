# -*- coding: utf-8 -*-
"""Which Unicode does Word's PDF give each Symbol-font private-use character?

The CLI PDF draws a Symbol bullet (U+F0B7) with the right glyph, but its ToUnicode
maps it back to U+F0B7 where Word's says U+2022, so copy-paste and text search differ
(zh dev FullPicture reports).  symbol.ttf carries no glyph names (post 3.0, uniF0xx), so
the table is read from Word: one paragraph per code point U+F020..U+F0FF in Symbol,
«[NNNN]» marker before each, Word PDF text extraction after the marker.

    python _pb_symbol_tounicode_gen.py gen / word
"""
import json
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_symbol_tounicode")
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for cp in range(0xF020, 0xF100):
        body.append('<w:p><w:r><w:t xml:space="preserve">[%04X]</w:t></w:r>'
                    '<w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol"/></w:rPr>'
                    '<w:sym w:font="Symbol" w:char="%04X"/></w:r>'
                    '<w:r><w:t xml:space="preserve">[END]</w:t></w:r></w:p>' % (cp, cp))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>'
           + "".join(body) + '<w:sectPr/></w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    with zipfile.ZipFile(os.path.join(OUT, "symbol.docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
    print("wrote", OUT)


def word():
    import fitz
    import win32com.client as w
    src = os.path.join(OUT, "symbol.docx")
    pdf = os.path.join(OUT, "symbol_word.pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(src, ReadOnly=True)
        d.ExportAsFixedFormat(pdf, 17)
        d.Close(False)
    finally:
        app.Quit()
    text = "".join(pg.get_text() for pg in fitz.open(pdf))
    table = {}
    for m in re.finditer(r"\[([0-9A-F]{4})\](.*?)\[END\]", text, re.S):
        got = m.group(2).strip()
        table[m.group(1)] = got
    json.dump(table, open(os.path.join(OUT, "word.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(len(table), {k: (v, [hex(ord(c)) for c in v]) for k, v in list(table.items()) if k in ("F0B7", "F061", "F0A7", "F020")})


if __name__ == "__main__":
    {"gen": gen, "word": word}[sys.argv[1]]()
