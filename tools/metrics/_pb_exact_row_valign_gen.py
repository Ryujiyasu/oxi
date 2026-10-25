# -*- coding: utf-8 -*-
"""In an EXACT-height row, what box does a vAlign=center cell centre in when another cell's
content is taller than the row?

ja/forms__02157d72 p1 row 2 (trHeight exact 323 = 16.15pt): the two-paragraph cells
«那覇市…» / «生年月日 / 月日» overflow the row; Word keeps «生年月日», «平熱» (vAlign center)
and the merged «現住所» where a one-line cell sits, Oxi puts them 8.07pt (= 323tw / 2) lower:
it centred in the content height (32pt) instead of the exact height.

Arms: row 1 exact 323 with [merged vAlign=center cell | a cell of two / one paragraph |
a one-line vAlign=center cell], row 2 exact 386 closing the merge; and the same with
atLeast. Prints each cell's first baseline, Word and Oxi.

    python tools/metrics/_pb_exact_row_valign_gen.py      (PB_OXI=<renderer.exe> adds Oxi)
"""
import json, os, subprocess, sys, tempfile, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_exact_row_valign")
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
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="320" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>') % W


def tc(w, paras, valign=None, vmerge=None):
    pr = '<w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s%s</w:tcPr>' % (
        w, '<w:vMerge w:val="restart"/>' if vmerge == "restart" else ('<w:vMerge/>' if vmerge == "cont" else ""),
        '<w:vAlign w:val="%s"/>' % valign if valign else "")
    return "<w:tc>" + pr + "".join('<w:p><w:r><w:t>%s</w:t></w:r></w:p>' % t for t in paras) + "</w:tc>"


def tr(h, cells, rule="exact"):
    return '<w:tr><w:trPr><w:trHeight w:hRule="%s" w:val="%d"/></w:trPr>%s</w:tr>' % (rule, h, "".join(cells))


def build(name, rows):
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="7350" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>'
           '<w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="1350"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
           + rows + "</w:tbl>")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body><w:p><w:r><w:t>前</w:t></w:r></w:p>%s<w:p><w:r><w:t>後</w:t></w:r></w:p>'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>') % (W, tbl)
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, name + ".docx")
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES); z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return p


def lines_word(pdf):
    import pymupdf
    ys = {}
    for b in pymupdf.open(pdf)[0].get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                t = "".join(c["c"] for c in s["chars"]).strip()
                if t:
                    ys.setdefault(t[:2], round(s["chars"][0]["origin"][1], 2))
    return ys


def lines_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    ys = {}
    for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"]:
        if c["char"].strip():
            ys.setdefault(c["char"], round(c["baseline"], 1))  # per character: cells share baselines
    return ys


if __name__ == "__main__":
    import win32com.client
    two = tc(3000, ["二行目あり", "二行目"]); one = tc(3000, ["一行だけ"])
    left = tc(1350, ["左"], "center", "restart"); mid = tc(3000, ["中央"], "center")
    second = tr(386, [tc(1350, [""], "center", "cont"), tc(3000, ["次行"]), tc(3000, ["次行"])])
    # a merged cell whose content (three 11pt lines, 48pt) overflows the two exact rows it spans (16.15 + 19.3)
    over = tc(1350, ["一段落目", "二段落目", "三段落目"], None, "restart")
    arms = [("exact_two", tr(323, [left, two, mid]) + second), ("exact_one", tr(323, [left, one, mid]) + second),
            ("atLeast_two", tr(323, [left, two, mid], "atLeast") + second),
            ("exact_vmerge_over", tr(323, [over, one, mid]) + second)]
    for name, rows in arms:
        p = build(name, rows); pdf = p[:-5] + ".pdf"
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
        finally:
            w.Quit()
        wy = lines_word(pdf)
        oy = lines_oxi(os.environ["PB_OXI"], p) if os.environ.get("PB_OXI") else {}
        print("%-12s " % name + "  ".join("%s W %.1f O %s" % (k, wy.get(k, 0), ("%.1f" % oy[k[0]]) if k[0] in oy else "-")
                                           for k in ("左", "二行", "一行", "中央", "次行", "後") if k in wy), flush=True)
