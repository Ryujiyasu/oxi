# -*- coding: utf-8 -*-
"""Which face does Word draw CJK text in when the run's eastAsia face is a Latin face?

zh dev educational__005919f8 (and five more FullPicture.app reports): docDefaults set
every rFonts slot to Calibri and only w:lang w:val="en-US"; the Chinese text has no
glyphs in Calibri.  Word's PDF draws it in PMingLiU (新細明體) at a line pitch of 18.84;
Oxi falls back to SimSun (18.0) and every CJK line drifts up 0.8pt.  Arms (one fresh
Word process per document, see the substitute-cache trap): eastAsia face {Calibri,
Arial, Times New Roman} x eastAsia lang {absent, zh-CN, zh-TW, ja-JP} x the theme
(the specimen's minorFont: Hans 宋体 / Hant 新細明體 / Jpan ＭＳ 明朝; or no theme).
Read: the PDF font of the CJK span and the line pitch of two CJK lines.

    python _pb_ea_latin_face_fallback_gen.py gen
    python _pb_ea_latin_face_fallback_gen.py word
    python _pb_ea_latin_face_fallback_gen.py oxi EXE   # needs the CLI (docx-to-pdf)
"""
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_ea_latin_face_fallback")
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
FACES = ["Calibri", "Arial", "Times New Roman"] + (os.environ.get("PB_FACES", "").split(",") if os.environ.get("PB_FACES") else [])
LANGS = [None, "zh-CN", "zh-TW", "ja-JP"]
THEMES = [True, False]
TEXT = "报告讲座主题为连续谱中的束缚态，涉及量子力学、电磁和声学谐振器中的束缚态和非束缚态。报告人研究领域包括连续谱中的束缚态、光子晶体波导传输等。"
THEME = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>'
         '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
         '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1>'
         '<a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>'
         '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink>'
         '<a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>'
         '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
         '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>'
         '<a:font script="Jpan" typeface="ＭＳ 明朝"/><a:font script="Hang" typeface="맑은 고딕"/>'
         '<a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/></a:minorFont></a:fontScheme>'
         '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
         '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
         '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
         '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
         '</a:fmtScheme></a:themeElements></a:theme>')


def arms():
    return [(f, l, t) for f in FACES for l in LANGS for t in THEMES]


def name(f, l, t):
    return "%s_%s_%s" % (f.replace(" ", ""), l or "nolang", "theme" if t else "notheme")


def gen():
    os.makedirs(OUT, exist_ok=True)
    for f, l, t in arms():
        lang = '<w:lang w:val="en-US"%s/>' % (' w:eastAsia="%s"' % l if l else "")
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
                  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
                  '<w:sz w:val="24"/><w:szCs w:val="24"/>%s</w:rPr></w:rPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
                  % (f, f, f, f, lang))
        body = '<w:p><w:r><w:t>%s</w:t></w:r></w:p><w:p><w:r><w:t>END</w:t></w:r></w:p>' % TEXT
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>' + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
               '</w:body></w:document>')
        ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
              '<Default Extension="xml" ContentType="application/xml"/>'
              '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
              '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
              + ('<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' if t else "")
              + '</Types>')
        drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                 + ('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' if t else "")
                 + '</Relationships>')
        with zipfile.ZipFile(os.path.join(OUT, name(f, l, t) + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", styles)
            if t:
                z.writestr("word/theme/theme1.xml", THEME)
    print("wrote", len(arms()))


def read_pdf(pdf):
    import fitz
    pg = fitz.open(pdf)[0]
    lines = []
    for b in pg.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            sp = [s for s in ln["spans"] if s["text"].strip()]
            if sp and any(ord(c) > 0x2E80 for c in sp[0]["text"]):
                lines.append((round(sp[0]["origin"][1], 2), sp[0]["font"]))
    lines.sort()
    pitch = round(lines[1][0] - lines[0][0], 2) if len(lines) > 1 else None
    return {"font": lines[0][1] if lines else None, "pitch": pitch}


def word():
    import win32com.client as w
    res = {}
    only = sys.argv[2:]
    for f, l, t in arms():
        n = name(f, l, t)
        if only and n not in only:
            continue
        if os.path.exists(os.path.join(OUT, n + "_word.pdf")):
            res[n] = read_pdf(os.path.join(OUT, n + "_word.pdf"))
            continue
        app = w.DispatchEx("Word.Application")  # fresh process: substitutes are cached per process
        app.Visible = False
        try:
            d = app.Documents.Open(os.path.join(OUT, n + ".docx"), ReadOnly=True)
            d.ExportAsFixedFormat(os.path.join(OUT, n + "_word.pdf"), 17)
            d.Close(False)
        finally:
            app.Quit()
        res[n] = read_pdf(os.path.join(OUT, n + "_word.pdf"))
        print(n, res[n], flush=True)
    json.dump(res, open(os.path.join(OUT, "word.json"), "w", encoding="utf-8"), indent=1)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json"), encoding="utf-8"))
    for f, l, t in arms():
        n = name(f, l, t)
        pdf = os.path.join(OUT, n + "_oxi.pdf")
        subprocess.run([exe, "docx-to-pdf", os.path.join(OUT, n + ".docx"), pdf], capture_output=True)
        got = read_pdf(pdf)
        flag = "OK " if got["pitch"] == want[n]["pitch"] else "DIFF"
        print(flag, n, "word", want[n], "oxi", got)


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
