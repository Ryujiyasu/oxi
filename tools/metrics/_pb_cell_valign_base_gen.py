# -*- coding: utf-8 -*-
"""Where does Word put the text baseline inside a table cell, relative to the
cell's top rule?  blind-G ja/forms__020115bc and administrative__19bb32f7: the
row rules match Oxi within 0.3pt but every cell line sits 1.2pt higher in Oxi
(MS Mincho 11pt, vAlign center, 18pt rows on a linesAndChars 350 grid: Word's
baseline is 13.08 below the rule, Oxi's 11.88).  Arms: one single-cell table
(width 4000) with one line, tblBorders single 4, default cell margins; variants
over vAlign x size x grid x row height:
    c11_g        center, 11pt, grid 350              t11_g   top      b11_g  bottom
    c11_n        center, 11pt, no grid               t11_n            b11_n
    c105_g / c12_g / c14_g   center at 10.5 / 12 / 14pt, grid 350
    c11_g_h720   center, 11pt, grid, trHeight 720 atLeast     t11_g_h720 / b11_g_h720
    c11_n_h720   center, 11pt, no grid, trHeight 720
    c11_g_cal    center, 11pt, grid, ascii Calibri + eastAsia MS Mincho, Latin text
Read: the cell's top and bottom rules (PDF drawings; Oxi border elements) and the
baseline, as offsets from the top rule.

    python tools/metrics/_pb_cell_valign_base_gen.py gen
    python tools/metrics/_pb_cell_valign_base_gen.py word
    python tools/metrics/_pb_cell_valign_base_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_cell_valign_base"
MS = "ＭＳ 明朝"  # MS Mincho (full-width)
# (valign, sz half-points, grid, trHeight twips or 0, latin)
ARMS = {"c11_g": ("center", 22, True, 0, False), "t11_g": ("top", 22, True, 0, False), "b11_g": ("bottom", 22, True, 0, False),
        "c11_n": ("center", 22, False, 0, False), "t11_n": ("top", 22, False, 0, False), "b11_n": ("bottom", 22, False, 0, False),
        "c105_g": ("center", 21, True, 0, False), "c12_g": ("center", 24, True, 0, False), "c14_g": ("center", 28, True, 0, False),
        "c11_g_h720": ("center", 22, True, 720, False), "t11_g_h720": ("top", 22, True, 720, False), "b11_g_h720": ("bottom", 22, True, 720, False),
        "c11_n_h720": ("center", 22, False, 720, False), "c11_g_cal": ("center", 22, True, 0, True),
        # the form's compat: adjustLineHeightInTable (cell lines snap to the grid pitch; 020115bc rows are 18.0)
        "c11_g_alh": ("center", 22, True, 0, False), "t11_g_alh": ("top", 22, True, 0, False), "b11_g_alh": ("bottom", 22, True, 0, False),
        "c11_g_h720_alh": ("center", 22, True, 720, False), "t11_g_h720_alh": ("top", 22, True, 720, False), "b11_g_h720_alh": ("bottom", 22, True, 720, False),
        "c12_g_alh": ("center", 24, True, 0, False), "c105_g_alh": ("center", 21, True, 0, False), "c11_n_alh": ("center", 22, False, 0, False)}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (va, sz, grid, trh, latin) in ARMS.items():
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="' + MS + '" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
                  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        fonts = ('<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="' + MS + '"/>') if latin else ('<w:rFonts w:ascii="' + MS + '" w:hAnsi="Arial" w:eastAsia="' + MS + '"/>')
        text = "Year 2024" if latin else "年月日第号"
        trpr = ('<w:trPr><w:trHeight w:val="%d" w:hRule="atLeast"/></w:trPr>' % trh) if trh else ''
        tbl = ('<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblBorders>'
               '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
               '<w:tr>' + trpr + '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vAlign w:val="' + va + '"/></w:tcPr>'
               '<w:p><w:pPr><w:jc w:val="center"/><w:rPr>' + fonts + '<w:kern w:val="0"/><w:sz w:val="%d"/></w:rPr></w:pPr>' % sz +
               '<w:r><w:rPr>' + fonts + '<w:kern w:val="0"/><w:sz w:val="%d"/></w:rPr><w:t>' % sz + text + '</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
        dg = '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="1228"/>' if grid else ''
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body><w:p><w:r><w:t>head</w:t></w:r></w:p>' + tbl +
               '<w:p><w:r><w:t>end</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' + dg + '</w:sectPr></w:body></w:document>')
        alh = '<w:adjustLineHeightInTable/>' if arm.endswith("_alh") else ''
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat>' + alh + '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / (arm + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def word():
    import win32com.client, fitz
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / (arm + ".docx"), t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / (arm + ".pdf")).resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in ARMS:
        pg = fitz.open(str(OUT / (arm + ".pdf")))[0]; rules = set(); base = None
        for dr in pg.get_drawings():
            for it in dr["items"]:
                if it[0] == "l" and abs(it[1].y - it[2].y) < 0.3 and abs(it[1].x - it[2].x) > 20: rules.add(round(it[1].y, 2))
                elif it[0] == "re" and it[1].height < 1.5 and it[1].width > 20: rules.add(round(it[1].y0 + it[1].height / 2, 2))
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip() and s["text"].strip() not in ("head", "end"): base = round(s["origin"][1], 2)
        rules = sorted(rules); top, bot = rules[0], rules[-1]
        res[arm] = (top, bot, base)
        print("%-12s word top=%7.2f bottom=%7.2f h=%6.2f base-top=%6.2f bottom-base=%6.2f" % (arm, top, bot, bot - top, base - top, bot - base))
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    for arm in ARMS:
        t = tempfile.mkdtemp(); lo = os.path.join(t, "l.json"); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / (arm + ".docx")), os.path.join(t, "p"), "--dump-layout=" + lo], capture_output=True, timeout=600)
        subprocess.run([os.path.abspath(exe), str(OUT / (arm + ".docx")), os.path.join(t, "q"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        els = json.load(open(lo, encoding="utf-8"))["pages"][0]["elements"]
        rules = sorted({round(e["y"], 2) for e in els if e.get("type") == "border" and e.get("w", 0) > 20 and e.get("h", 0) < 1.5})
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]
        top, bot = rules[0], rules[-1]
        base = max(round(x["baseline"], 2) for x in g if x["char"].strip() and top < x["baseline"] < bot + 2)
        print("%-12s oxi  top=%7.2f bottom=%7.2f h=%6.2f base-top=%6.2f bottom-base=%6.2f" % (arm, top, bot, bot - top, base - top, bot - base))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
