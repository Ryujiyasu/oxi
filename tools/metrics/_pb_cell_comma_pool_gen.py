# -*- coding: utf-8 -*-
"""How much can a justified compat-15 line squeeze its 、 marks -- in a cell and in the body?

golden probethdr: a jc=both cell line of 30 ideographs + 4 「、」 (357pt natural) holds in
~349.3pt in Word (each 、 drawn 8.5 wide); Oxi's cell breaker (S1174: a half-em pool per
line) wraps the last ideograph. The same paragraph in the body already matches Word.

Arms: container (cell / body) x number of 、 among the first 20 characters (0..6).
Per arm: binary-search the narrowest width (5tw steps) at which the first line still
holds all 20 characters; shortfall = natural width - that width.

    python tools/metrics/_pb_cell_comma_pool_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cell_comma_pool")
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
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>' % W)
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>' % W)
KANJI = "本項定事項関係法令及規程趣旨照善良管理者注意誠実適切取扱処理"
TAIL = "以下本文続行末尾確認用文字列以下本文続行末尾確認用文字列"
FS = 10.5


def line(k):
    """20 characters: k 「、」 spread among ideographs, never first or last."""
    chars = list(KANJI[:20 - k])
    for i in range(k):
        chars.insert(2 + i * 3, "、")
    return "".join(chars)


def natural(k):
    return 20 * FS


def build(cont, k, width_tw):
    text = line(k) + TAIL
    para = '<w:p><w:pPr><w:jc w:val="both"/>%s</w:pPr><w:r><w:t>%s</w:t></w:r></w:p>'
    if cont == "cell":
        body = ('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl><w:p/>'
                % (width_tw, width_tw, width_tw, para % ("", text)))
    else:
        body = para % ('<w:ind w:left="0" w:right="%d"/>' % (9070 - width_tw), text)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>'
           '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>' % (W, body))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "cp_%s_k%d_w%d.docx" % (cont, k, width_tw))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return path


def first_line_count(w, path):
    d = w.Documents.Open(path, ReadOnly=True)
    try:
        r = d.Paragraphs(1).Range
        y0 = d.Range(r.Start, r.Start).Information(6)
        n = 0
        for i in range(1, 26):
            c = d.Range(r.Start + i - 1, r.Start + i - 1)
            if abs(c.Information(6) - y0) > 1:
                break
            n = i
        return n
    finally:
        d.Close(False)


if __name__ == "__main__":
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for cont in ("cell", "body"):
            for k in range(0, 7):
                lo, hi = 3600, 4200  # tw: lo fails (< 20 chars), hi holds
                while hi - lo > 5:
                    mid = (lo + hi) // 2 // 5 * 5
                    if mid in (lo, hi):
                        break
                    if first_line_count(w, build(cont, k, mid)) >= 20:
                        hi = mid
                    else:
                        lo = mid
                print("%s k=%d: narrowest holding 20 chars %d tw = %.2fpt, shortfall %.2fpt (%.3f em)"
                      % (cont, k, hi, hi / 20, natural(k) - hi / 20, (natural(k) - hi / 20) / FS), flush=True)
    finally:
        w.Quit()
