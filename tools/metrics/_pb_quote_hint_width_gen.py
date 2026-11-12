# -*- coding: utf-8 -*-
"""How wide does Word draw a curly quote in its own run, by w:hint, language and faces?

zh dev correspondence__015abfaf: «现将2024 年度威海市“聚力抓经济 建功新时代”争先创…» at 16pt, the
quotes each in their own run (open: hint=default, lang en / eastAsia zh-CN; close: hint=eastAsia,
lang eastAsia zh-CN), ascii «Nimbus Roman» and eastAsia «CESI仿宋-GB2312» both missing.  Word's PDF
advances the open quote 6.6pt, the close 8.7; Oxi draws both full width (15.7) and the line holds
two characters fewer.  Arms: faces {TNR + SimSun, both missing (the specimen's names)} x hint
{none, default, eastAsia} x lang {none, en, en + eastAsia zh-CN}, 16pt, in «前“聚力”后».
Read: the advance of «“» (x of 聚 - x of “) and of «”» (x of 后 - x of ”) in Word's PDF.

    python _pb_quote_hint_width_gen.py gen / word / oxi EXE
"""
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_quote_hint_width")
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
FACES = {"real": ("Times New Roman", "SimSun"), "missing": ("Nimbus Roman", "CESI仿宋-GB2312")}
HINTS = [None, "default", "eastAsia"]
LANGS = {"none": None, "en": '<w:lang w:val="en"/>', "en_zh": '<w:lang w:val="en" w:eastAsia="zh-CN"/>'}


def arms():
    return [(f, h, l) for f in FACES for h in HINTS for l in LANGS]


def rpr(face, hint, lang):
    a, ea = FACES[face]
    h = ' w:hint="%s"' % hint if hint else ""
    return ('<w:rPr><w:rFonts%s w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/><w:sz w:val="32"/>%s</w:rPr>'
            % (h, a, a, ea, a, LANGS[lang] or ""))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for f, h, l in arms():
        base = rpr(f, None, "none")
        q = rpr(f, h, l)
        body.append('<w:p><w:r>%s<w:t>%s:</w:t></w:r><w:r>%s<w:t>前</w:t></w:r><w:r>%s<w:t>“</w:t></w:r>'
                    '<w:r>%s<w:t>聚力</w:t></w:r><w:r>%s<w:t>”</w:t></w:r><w:r>%s<w:t>后</w:t></w:r></w:p>'
                    % (rpr("real", None, "none"), "%s_%s_%s" % (f, h or "nohint", l), base, q, base, q, base))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
              '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
              'w:eastAsia="SimSun" w:cs="Times New Roman"/><w:sz w:val="32"/>'
              '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>'
           + "".join(body) + '<w:sectPr/></w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    with zipfile.ZipFile(os.path.join(OUT, "quote.docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)
    print("wrote", len(arms()), "arms")


def read(pdf):
    import fitz
    res = {}
    for pg in fitz.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                chars = [c for s in ln["spans"] for c in s["chars"]]
                txt = "".join(c["c"] for c in chars)
                if ":" not in txt:
                    continue
                label = txt.split(":")[0].strip()
                xs = {c["c"]: c["bbox"][0] for c in chars}
                try:
                    res[label] = (round(xs["聚"] - xs["“"], 2), round(xs["后"] - xs["”"], 2))
                except KeyError:
                    res[label] = None
    return res


def word():
    import win32com.client as w
    src = os.path.join(OUT, "quote.docx")
    pdf = os.path.join(OUT, "quote_word.pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(src, ReadOnly=True)
        d.ExportAsFixedFormat(pdf, 17)
        d.Close(False)
    finally:
        app.Quit()
    res = read(pdf)
    json.dump(res, open(os.path.join(OUT, "word.json"), "w"), indent=1)
    for k, v in res.items():
        print(k, v)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json")))
    pdf = os.path.join(OUT, "quote_oxi.pdf")
    subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, "quote.docx"), pdf], capture_output=True)
    got = read(pdf)
    for k in want:
        print("%-28s word %-16s oxi %s" % (k, want[k], got.get(k)))


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
