# -*- coding: utf-8 -*-
"""Where does the text that is pushed below a full-width FLOATING table start: at the
last row's bottom edge, or below the bottom border too?

ja/forms__02157d72 p1 carries three text-anchored full-width floating tables
(tblpY 80 / 360, tblOverlap never, borders sz 4). Word PDF: the paragraph after each
starts 0.43-0.9 below the bottom border's top edge (y0), Oxi starts it AT y0 (= its
last row's bottom). Inline tables (page.rs advances the bottom border width after the
last row) are the control.

Sheet: «前» (ＭＳ 明朝 10.5), a table of two exact 340tw rows (one short run each), «後»
(the anchor of the floating table). Arms: placement (float: tblpPr vertAnchor=text
horzAnchor=margin tblpY=80 tblOverlap=never / inline) x border size (4 / 12 / 24 / none).
Read: the baselines of «前», «後» and the table's horizontal rules (Word PDF drawings;
Oxi --dump-layout border elements), printed as «後» baseline minus the last rule's y0
and the rule positions themselves.

    python tools/metrics/_pb_floattbl_bottom_gen.py            (PB_OXI=<renderer.exe> adds Oxi; PB_GRID=lc|lines|none; PB_BORDERS=cell)
"""
import json, os, subprocess, sys, tempfile, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_floattbl_bottom")
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
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:compat><w:useFELayout/>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>') % W
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝"/><w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>') % W
GRID = {"lc": '<w:docGrid w:type="linesAndChars" w:linePitch="290" w:charSpace="-1209"/>', "lines": '<w:docGrid w:type="lines" w:linePitch="290"/>',
        "none": ""}[os.environ.get("PB_GRID", "lc")]  # PB_GRID=lc|lines|none
CELL = os.environ.get("PB_BORDERS") == "cell"
FLOAT = '<w:tblpPr w:leftFromText="142" w:rightFromText="142" w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="80"/><w:tblOverlap w:val="never"/>'


def build(name, placement, sz):
    borders = "" if sz is None else ('<w:tblBorders>' + "".join('<w:%s w:val="single" w:sz="%d" w:space="0" w:color="auto"/>' % (e, sz)
                                                                 for e in ("top", "left", "bottom", "right", "insideH", "insideV")) + '</w:tblBorders>')
    # PB_BORDERS=cell writes the same edges as tcBorders on every cell (no tblBorders), the
    # shape of ja/forms__02157d72's floats (no tblStyle, no tblBorders, every cell bordered)
    cell_borders = "" if sz is None or CELL else borders
    tc_borders = "" if sz is None or not CELL else borders.replace("tblBorders", "tcBorders")
    tbl = ('<w:tbl><w:tblPr>%s<w:tblW w:w="10000" w:type="dxa"/>%s<w:tblLayout w:type="fixed"/>'
           '<w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="10000"/></w:tblGrid>'
           % (FLOAT if placement == "float" else "", cell_borders))
    for t in ("一", "二"):
        tbl += ('<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="340"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="10000" w:type="dxa"/>%s</w:tcPr>'
                '<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr><w:r><w:t>%s</w:t></w:r></w:p></w:tc></w:tr>' % (tc_borders, t))
    tbl += "</w:tbl>"
    sp = '<w:pPr><w:snapToGrid w:val="0"/></w:pPr>'
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body><w:p>%s<w:r><w:t>前</w:t></w:r></w:p>%s<w:p>%s<w:r><w:t>後</w:t></w:r></w:p>'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="567" w:right="567" w:bottom="284" w:left="1134" w:header="851" w:footer="992"/>'
           '%s</w:sectPr></w:body></w:document>') % (W, sp, tbl, sp, GRID)
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, name + ".docx")
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES); z.writestr("word/settings.xml", SETTINGS); z.writestr("word/document.xml", doc)
    return p


def read_word(pdf):
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    ys = {}
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"] in "前後一二":
                        ys.setdefault(c["c"], c["origin"][1])
    rules = sorted({round(d["rect"].y0, 2) for d in pg.get_drawings() if d["rect"].width > 100 and d["rect"].height < 3})
    return ys, rules


def read_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json"); lay = os.path.join(t, "l.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    subprocess.run([exe, docx, os.path.join(t, "p2"), "--dump-layout=" + lay], capture_output=True)
    ys = {}
    for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"]:
        if c["char"] in "前後一二":
            ys.setdefault(c["char"], c["baseline"])
    rules = sorted({round(e["y"], 2) for e in json.load(open(lay, encoding="utf-8"))["pages"][0]["elements"]
                    if e.get("type") == "border" and e.get("w", 0) > 100})
    return ys, rules


def fmt(ys, rules):
    last = rules[-1] if rules else float("nan")
    return "前 %.2f 一 %.2f 後 %.2f  後-lastrule %.2f  rules %s" % (ys.get("前", 0), ys.get("一", 0), ys.get("後", 0), ys.get("後", 0) - last, rules)


if __name__ == "__main__":
    import win32com.client
    exe = os.environ.get("PB_OXI")
    for placement in ("float", "inline"):
        for sz in (4, 12, 24, None):
            name = "%s_sz%s_%s%s" % (placement, sz if sz else "none", os.environ.get("PB_GRID", "lc"), "_cell" if CELL else "")
            p = build(name, placement, sz); pdf = p[:-5] + ".pdf"
            w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
            try:
                d = w.Documents.Open(p, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
            finally:
                w.Quit()
            print("%-12s W " % name + fmt(*read_word(pdf)), flush=True)
            if exe:
                print("%-12s O " % "" + fmt(*read_oxi(exe, p)), flush=True)
