# -*- coding: utf-8 -*-
"""Where does a tab go when the text before it ends EXACTLY on a stop?

ja/educational__08709ff2: numbering «⑤» (BIZ UDPGothic 18pt, advance 18.00) with
w:ind left=360 hanging=360 -- the marker ends exactly on the hanging-indent stop.
Word draws the text at +21pt (the next 420tw default stop), not at +18; Oxi at +18.

Arms (one paragraph each, defaultTabStop 420):
  num_eq / num_lt / num_gt  -- numbering marker 18pt wide, hanging 360 / 400 / 320
  num_wide / num_wide_eq    -- hanging 900 / 840: a default stop (420) lies between
                               the marker end and the indent
  body_eq / body_lt         -- body text «亜» (18pt) then a tab, custom stop at 360 / 400
  body_eq_def               -- body «亜亜» (36pt) then a tab, no custom stop (default
                               stops 420 / 840: 36pt < 42 -> 42)
  body_def_exact            -- body «亜亜亜亜亜亜亜» 7x18=126pt... default stop 2520/20=126
Per arm: x of the first character after the tab, relative to the left margin.

    python tools/metrics/_pb_tab_exact_stop_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_tab_exact_stop")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
FACE = os.environ.get("PB_FACE", "BIZ UDPGothic")
# PB_FLAG=1: <w:doNotUseIndentAsNumberingTabStop/> in w:compat (08709ff2 carries it)
FLAG = os.environ.get("PB_FLAG") == "1"
RPR = '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="36"/></w:rPr>' % (FACE, FACE, FACE)
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
NUM = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering %s><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="5"/>'
       '<w:numFmt w:val="decimalEnclosedCircle"/><w:lvlText w:val="%%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl>'
       '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>' % W)


def para(name):
    if name.startswith("num_"):
        h = {"num_eq": 360, "num_lt": 400, "num_gt": 320, "num_wide": 900, "num_wide_eq": 840}[name]
        return ('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="%d" w:hanging="%d"/>%s</w:pPr>'
                '<w:r>%s<w:t>検索結果</w:t></w:r></w:p>' % (h, h, RPR, RPR))
    text, stop = {"body_eq": ("亜", 360), "body_lt": ("亜", 400), "body_eq_def": ("亜亜", None),
                  "body_def_exact": ("亜亜亜亜亜亜亜", None)}[name]
    tabs = '<w:tabs><w:tab w:val="left" w:pos="%d"/></w:tabs>' % stop if stop else ""
    return ('<w:p><w:pPr>%s%s</w:pPr><w:r>%s<w:t>%s</w:t></w:r><w:r>%s<w:tab/></w:r><w:r>%s<w:t>検索</w:t></w:r></w:p>'
            % (tabs, RPR, RPR, text, RPR, RPR))


ARMS = ["num_eq", "num_lt", "num_gt", "num_wide", "num_wide_eq", "body_eq", "body_lt", "body_eq_def", "body_def_exact"]


def build(cm):
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1701" w:right="1701"/></w:sectPr>'
           '</w:body></w:document>' % (W, "".join(para(a) for a in ARMS)))
    flag = '<w:doNotUseIndentAsNumberingTabStop/>' if FLAG else ''
    st = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:defaultTabStop w:val="420"/><w:compat>%s'
          '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>' % (W, flag, cm))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "tabx_%s_c%d%s.docx" % (FACE.replace(" ", ""), cm, "_flag" if FLAG else ""))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/settings.xml", st)
        z.writestr("word/numbering.xml", NUM); z.writestr("word/document.xml", doc)
    return path


def after_tab_x(pdf):
    import fitz
    out = []
    for b in fitz.open(pdf)[0].get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            ch = sorted([c for s in l["spans"] for c in s["chars"] if c["c"].strip()], key=lambda c: c["origin"][0])
            k = next((i for i, c in enumerate(ch) if c["c"] == "検"), None)
            if k is not None:
                out.append(round(ch[k]["origin"][0] - 85.05, 2))
    return out


if __name__ == "__main__":
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for cm in (11, 12, 14, 15):
            docx = build(cm); pdf = docx[:-5] + ".pdf"
            d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
            print("compat %d %s%s  %s" % (cm, FACE, " +flag" if FLAG else "", "  ".join("%s=%s" % a for a in zip(ARMS, after_tab_x(pdf)))), flush=True)
    finally:
        w.Quit()
