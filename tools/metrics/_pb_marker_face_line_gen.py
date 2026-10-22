# -*- coding: utf-8 -*-
"""How tall is a numbered line whose marker face differs from the text face?

en/policies__0097fbf2: Calibri 10 list items whose numbering marker is Arial Unicode MS
(fontTable charset 80 -> Word draws it in MS-Mincho).  Word's numbered lines sit ~0.5pt
lower than Oxi's per item; S1614 (AUM's own 1.332em box) over-corrected by 4pt.

Arms: marker face (MS Mincho / AUM declared charset 80 / AUM undeclared / Arial / same as
text) x text size (10 / 11) x line rule (single / 1.15).  Each doc: plain, numbered,
plain, numbered, plain ...; reads the PDF baselines and prints the numbered line's extra
advance over a plain line.

    python tools/metrics/_pb_marker_face_line_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_marker_face_line")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
         '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>')
FACES = {"mincho": "ＭＳ 明朝", "aum80": "Arial Unicode MS", "aum0": "Arial Unicode MS", "arial": "Arial", "same": "Calibri"}


def build(face, sz, line):
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri"/><w:sz w:val="%d"/><w:lang w:val="en-US" w:eastAsia="en-US"/>'
              '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % (W, sz, line)
    f = FACES[face]
    numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering %s><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
                 '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%%1"/><w:lvlJc w:val="left"/>'
                 '<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/></w:rPr></w:lvl></w:abstractNum>'
                 '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>') % (W, f, f, f)
    charset = {"aum80": '<w:charset w:val="80"/>', "aum0": '<w:charset w:val="00"/>'}.get(face, "")
    fonttable = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts %s>'
                 '<w:font w:name="Arial Unicode MS">%s<w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font></w:fonts>') % (W, charset)
    body = ""
    for i in range(1, 13):
        body += '<w:p><w:r><w:t>plain line %d text here</w:t></w:r></w:p>' % i
        body += '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>numbered item %d text</w:t></w:r></w:p>' % i
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>'
           % (W, R, body))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "mf_%s_%d_%d.docx" % (face, sz, line))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", styles)
        z.writestr("word/numbering.xml", numbering); z.writestr("word/fontTable.xml", fonttable)
        z.writestr("word/document.xml", doc)
    return path


def baselines(pdf):
    import fitz
    out = []
    for b in fitz.open(pdf)[0].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"])
            if "line" in t or "item" in t:
                out.append((round(l["spans"][-1]["origin"][1], 3), t.strip()[:12]))
    return sorted(set(out))


if __name__ == "__main__":
    import win32com.client
    for face in FACES:
        for sz in (20, 22):
            for line in (240, 276):
                p = build(face, sz, line); pdf = p[:-5] + ".pdf"
                w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
                try:
                    d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
                finally:
                    w.Quit()
                b = baselines(pdf)
                steps = [round(y2 - y1, 3) for (y1, _), (y2, _) in zip(b, b[1:])]
                plain_to_num = [s for k, s in enumerate(steps) if k % 2 == 0][:4]
                num_to_plain = [s for k, s in enumerate(steps) if k % 2 == 1][:4]
                print("%-6s sz%d line%d  plain->num %s  num->plain %s" % (face, sz, line, plain_to_num, num_to_plain), flush=True)
