# -*- coding: utf-8 -*-
"""Does compressPunctuation squeeze brackets to fit one more character?  By language?

zh dev legal__00558607 (compressPunctuation, eastAsia zh-CN, YaHei 16pt, left-aligned):
Word ends the line «…山东省人民政府办公» and compresses only the 》（ pair; Oxi squeezes
〔 〕 ） 《 by 3.2pt each and fits «厅» too.  Same class: policies__0060058e,
legal__005764f9.  Arms: eastAsia face (Microsoft YaHei, SimSun, MS Mincho) x lang
(zh-CN, ja-JP) x right indent 0..30pt in 1.5pt steps, the real line's text (16pt,
left-aligned, compressPunctuation in settings).  Read: characters on line 1.

    python _pb_cp_fit_lang_gen.py gen
    python _pb_cp_fit_lang_gen.py word
    python _pb_cp_fit_lang_gen.py oxi EXE
"""
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cp_fit_lang" + os.environ.get("PB_COMPAT", "") + os.environ.get("PB_JC", ""))
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + '>'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            + ("" if os.environ.get("PB_COMPAT") == "none" else
               '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%s"/></w:compat>'
               % os.environ.get("PB_COMPAT", "15")) +
            '</w:settings>')
TEXT = "导意见》（国办发〔2013〕15 号）和《山东省人民政府办公厅关于进一步加强"
FACES = ["Microsoft YaHei", "SimSun", "MS Mincho"]
LANGS = ["zh-CN", "ja-JP"]
INDENTS = [round(k * 1.5, 1) for k in range(21)]


def arms():
    return [(f, l, ind) for f in FACES for l in LANGS for ind in INDENTS]


def docx(face, lang):
    return os.path.join(OUT, "cp_%s_%s.docx" % (face.replace(" ", ""), lang))


def gen():
    os.makedirs(OUT, exist_ok=True)
    for face in FACES:
        for lang in LANGS:
            rpr = ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="%s"/>'
                   '<w:sz w:val="32"/><w:lang w:val="en-US" w:eastAsia="%s"/></w:rPr>' % (face, lang))
            body = []
            for ind in INDENTS:
                body.append('<w:p><w:pPr><w:ind w:right="%d"/><w:jc w:val="%s"/>%s</w:pPr>'
                            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                            % (int(round(ind * 20)), os.environ.get("PB_JC", "left"), rpr, rpr, TEXT))
            styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
                      '<w:docDefaults><w:rPrDefault>%s</w:rPrDefault>'
                      '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
                      '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>'
                      '</w:styles>' % rpr)
            doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>'
                   + "".join(body) + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                   '<w:pgMar w:top="1000" w:right="1814" w:bottom="1000" w:left="1800" w:header="500" w:footer="500" w:gutter="0"/>'
                   '</w:sectPr></w:body></w:document>')
            with zipfile.ZipFile(docx(face, lang), "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("[Content_Types].xml", CT)
                z.writestr("_rels/.rels", RELS)
                z.writestr("word/_rels/document.xml.rels", DRELS)
                z.writestr("word/document.xml", doc)
                z.writestr("word/styles.xml", styles)
                z.writestr("word/settings.xml", SETTINGS)
    print("wrote", len(FACES) * len(LANGS), "docs x", len(INDENTS), "indents")


def word():
    """Characters on each paragraph's first line, from the y of every character."""
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    res = {}
    try:
        for face in FACES:
            for lang in LANGS:
                d = app.Documents.Open(docx(face, lang), ReadOnly=True)
                counts = []
                for p in d.Paragraphs:
                    r = p.Range
                    y0 = d.Range(r.Start, r.Start).Information(6)
                    n = 0
                    for i in range(r.Start, r.End - 1):
                        if d.Range(i, i).Information(6) != y0:
                            break
                        n += 1
                    counts.append(n)
                d.Close(False)
                res["%s|%s" % (face, lang)] = counts
    finally:
        app.Quit()
    json.dump(res, open(os.path.join(OUT, "word.json"), "w", encoding="utf-8"), indent=1)
    for k, v in res.items():
        print(k, v)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json"), encoding="utf-8"))
    bad = 0
    for face in FACES:
        for lang in LANGS:
            dump = docx(face, lang)[:-5] + "_oxi.json"
            subprocess.run([exe, docx(face, lang), dump[:-5], "--dump-layout=" + dump], capture_output=True)
            d = json.load(open(dump, encoding="utf-8"))
            first = {}
            for pg in d["pages"]:
                for e in pg["elements"]:
                    if e["type"] != "text" or e.get("para_idx") is None:
                        continue
                    pi = e["para_idx"]
                    y = round(e["y"], 1)
                    if pi not in first:
                        first[pi] = [y, 0]
                    if y == first[pi][0]:
                        first[pi][1] += len((e.get("text") or "").replace(" ", "")) + (e.get("text") or "").count(" ")
            got = [first[k][1] for k in sorted(first)]
            key = "%s|%s" % (face, lang)
            print(key, "word", want[key])
            print(key.rjust(len(key)), "oxi ", got)
            bad += sum(1 for a, b in zip(want[key], got) if a != b)
    print("mismatching arms", bad)


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
