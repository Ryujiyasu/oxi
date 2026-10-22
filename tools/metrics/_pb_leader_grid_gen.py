# -*- coding: utf-8 -*-
"""Where does Word put the dots of a tab leader?

ja/policies__1a7a3fec (TOC, ＭＳ ゴシック 11pt, right tab with dot leader): Word draws 52 dots at
x = k x 5.64 (page-absolute, k = 34..85), the last ending before the page number; Oxi
draws floor(gap / 5.5) = 53 and the renderer's advance pushes the last onto the number.

Arms: face x size x left indent of the text before the tab. Each paragraph: «Item» + tab
(right stop with dot leader at 8000tw) + «12». Prints the dot count, the step, and the
first dot's x divided by the step (integer = page-absolute grid).

    python tools/metrics/_pb_leader_grid_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_leader_grid")
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
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
FACES = {"gothic": "ＭＳ ゴシック", "mincho": "ＭＳ 明朝", "calibri": "Calibri", "tnr": "Times New Roman"}
SIZES = [18, 21, 22, 24]
INDENTS = [0, 100, 250]
BOLD = os.environ.get("PB_BOLD") == "1"  # bold runs (ＭＳ faces have no bold: synthetic)


def build():
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % W
    body = ""
    for face, f in FACES.items():
        for sz in SIZES:
            for ind in INDENTS:
                rpr = '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/>%s</w:rPr>' % (f, f, f, sz, "<w:b/>" if BOLD else "")
                body += ('<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8000"/></w:tabs><w:ind w:left="%d"/></w:pPr>'
                         '<w:r>%s<w:t>Item%s%d%d</w:t></w:r><w:r>%s<w:tab/></w:r><w:r>%s<w:t>12</w:t></w:r></w:p>'
                         % (ind, rpr, face[0].upper(), sz, ind, rpr, rpr))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440"/></w:sectPr></w:body></w:document>' % (W, body))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "leader%s.docx" % ("_b" if BOLD else ""))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    import win32com.client, pymupdf
    p = build(); pdf = p[:-5] + ".pdf"
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    for pg in pymupdf.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                ch = [(c["c"], c["origin"][0], c["bbox"][2], s["size"]) for s in l["spans"] for c in s["chars"]]
                t = "".join(c[0] for c in ch)
                if not t.startswith("Item"):
                    continue
                dots = [c for c in ch if c[0] == "."]
                tag = t[4:].split(".")[0].strip()
                num = [c for c in ch if c[0] in "12"][-2]
                if len(dots) < 2:
                    print(tag, "no dots"); continue
                step = (dots[-1][1] - dots[0][1]) / (len(dots) - 1)
                print("%-8s n=%3d step=%.3f first=%.2f first/step=%.3f lastend=%.2f num=%.2f dotsize=%.2f"
                      % (tag, len(dots), step, dots[0][1], dots[0][1] / step, dots[-1][2], num[1], dots[0][3]), flush=True)
