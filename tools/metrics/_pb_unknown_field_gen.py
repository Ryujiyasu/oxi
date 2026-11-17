# -*- coding: utf-8 -*-
"""A complex field whose instruction starts with a word Word does not know as a
field name: Word reads the word as a bookmark name (a bare-bookmark REF). The
docx corpus replaces instructions with «sanitize…», so its result-less fields
are all of this kind (ja/reference__0b6f3b32 / 13e1b7fc / 530de8d7: one
«エラー! ブックマークが定義されていません。» per result-less field in the Word PDF).

    python _pb_unknown_field_gen.py gen OUTDIR
    python _pb_unknown_field_gen.py pdf OUTDIR    # Word PDF text per arm
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="18"/></w:rPr>'


def fld(instr, result=None):
    r = lambda inner: "<w:r>" + RPR + inner + "</w:r>"
    out = r('<w:fldChar w:fldCharType="begin"/>') + r('<w:instrText xml:space="preserve">%s</w:instrText>' % instr)
    if result is not None:
        out += r('<w:fldChar w:fldCharType="separate"/>') + r("<w:t>%s</w:t>" % result)
    return out + r('<w:fldChar w:fldCharType="end"/>')


ARMS = [
    ("unknown_nores", fld("sanitizesanitize")),
    ("unknown_res", fld("sanitizesanitize", "キャッシュ")),
    ("author_nores", fld(" AUTHOR ")),
    ("bookmark_nores", fld(" mark1 ")),
]


def docx(out, label):
    return os.path.join(out, "uf_%s.docx" % label)


def gen(out):
    os.makedirs(out, exist_ok=True)
    for label, body in ARMS:
        para = ('<w:p><w:r>' + RPR + '<w:t>前</w:t></w:r>' + body + '<w:r>' + RPR + '<w:t>後</w:t></w:r></w:p>'
                '<w:p><w:bookmarkStart w:id="0" w:name="mark1"/><w:r>' + RPR + '<w:t>しおり</w:t></w:r>'
                '<w:bookmarkEnd w:id="0"/></w:p>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + "><w:body>"
               + para + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" '
               'w:bottom="1134" w:left="1134"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(docx(out, label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/document.xml", doc)


def pdf(out):
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for label, _ in ARMS:
            src = docx(out, label)
            dst = src[:-5] + ".pdf"
            d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            try:
                d.ExportAsFixedFormat(dst, 17)
            finally:
                d.Close(False)
            line = fitz.open(dst)[0].get_text().splitlines()[0]
            print("%-16s %r" % (label, line))
    finally:
        app.Quit()


if __name__ == "__main__":
    {"gen": gen, "pdf": pdf}[sys.argv[1]](sys.argv[2])
