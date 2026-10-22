# -*- coding: utf-8 -*-
"""Does a Latin word in a CJK-face run wrap whole in a table cell?

ja/reference__3187a5b4: a 2122tw cell holds «(physical condition)» in ＭＳ Ｐゴシック 12pt (compat 15,
compressPunctuation, balanceSingleByteDoubleByteWidth). Word wraps at the space
(«(physical» / «condition)»); Oxi broke inside the word («(physical conditio» / «n)»).

Arms: cell vs body (same width via right indent) x face (ＭＳ Ｐゴシック / ＭＳ ゴシック / Century)
x width sweep (1900..2300tw). Prints, per arm and width, line 1 of the text.

    python tools/metrics/_pb_cell_latin_wordwrap_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cell_latin_wordwrap")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:balanceSingleByteDoubleByteWidth/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
FACES = {"pgothic": "ＭＳ Ｐゴシック", "gothic": "ＭＳ ゴシック", "century": "Century"}
TEXT = "(physical condition)"
WIDTHS = list(range(1700, 2400, 20))


def build(face):
    f = FACES[face]
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:kern w:val="0"/><w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
              '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % (W, f, f, f)
    body = ""
    for w in WIDTHS:
        body += ('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
                 '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>%s</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
                 '<w:p><w:r><w:t>C%d</w:t></w:r></w:p>' % (w, w, w, TEXT, w))
        # body twin: text width = cell width - 2 x 108tw default margins
        body += ('<w:p><w:pPr><w:ind w:right="%d"/></w:pPr><w:r><w:t>%s</w:t></w:r></w:p><w:p><w:r><w:t>B%d</w:t></w:r></w:p>'
                 % (9026 - (w - 216), TEXT, w))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/></w:sectPr></w:body></w:document>' % (W, body))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "lw_%s.docx" % face)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return path


def first_lines(lines):
    """[(text)] in reading order -> {('C'|'B', width): line-1 text}"""
    out, cur = {}, []
    for t in lines:
        t = t.strip()
        if len(t) > 1 and t[0] in "CB" and t[1:].isdigit():
            out[(t[0], int(t[1:]))] = cur[0] if cur else ""
            cur = []
        elif t.strip():
            cur.append(t.strip())
    return out


def word_lines(pdf):
    import pymupdf
    res = []
    for pg in pymupdf.open(pdf):
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                res.append("".join(s["text"] for s in l["spans"]))
    return res


def oxi_lines(exe, docx):
    import subprocess, json, tempfile
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    res = []
    for p in json.load(open(g, encoding="utf-8"))["pages"]:
        rows = {}
        for c in p["glyphs"]:
            rows.setdefault(round(c["baseline"], 1), []).append(c)
        for k in sorted(rows):
            res.append("".join(c["char"] for c in sorted(rows[k], key=lambda c: c["x"])))
    return res


if __name__ == "__main__":
    import win32com.client
    exe = os.environ.get("PB_OXI")
    for face in FACES:
        p = build(face); pdf = p[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        wl = first_lines(word_lines(pdf))
        ol = first_lines(oxi_lines(exe, p)) if exe else {}
        for kind in "CB":
            seq = []
            for wd in WIDTHS:
                a = wl.get((kind, wd), "?").replace(" ", "")
                b = ol.get((kind, wd), "-").replace(" ", "")
                seq.append("%d:%s%s" % (wd, a[-6:], "" if not exe or a == b else "|O:" + b[-6:]))
            print(face, kind, " ".join(seq), flush=True)
