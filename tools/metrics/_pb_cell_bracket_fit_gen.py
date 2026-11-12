# -*- coding: utf-8 -*-
"""How many characters does a narrow table cell line hold around a full-width bracket?

zh dev forms__00243d36 p8 header: «权利人（标准起草单位）» in a 1098tw cell (SimSun 14pt, jc center,
exact 20pt, compressPunctuation, compat 14, eastAsia zh-CN): Word breaks «权利人 / （标准 / 起草单
/ 位）», Oxi «… / （标准起 / …» -- one more character, so the cell has fewer lines and its
vertically centred text sits a line lower.  Arms: lang {zh-CN, ja-JP} x compat {14, 15} x
cell width 1000..1400tw in 25tw steps (one row per width).  Read: characters per line, from
Word's per-character vertical position (COM).

    python _pb_cell_bracket_fit_gen.py gen / word / oxi EXE
"""
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cell_bracket_fit")
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TEXT = "权利人（标准起草单位）"
WIDTHS = list(range(1000, 1401, 25))
ARMS = [(l, c) for l in ("zh-CN", "ja-JP") for c in (14, 15)]


def name(lang, compat):
    return "cell_%s_c%d" % (lang, compat)


def gen():
    os.makedirs(OUT, exist_ok=True)
    for lang, compat in ARMS:
        rpr = ('<w:rPr><w:rFonts w:hint="eastAsia" w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体" w:cs="宋体"/>'
               '<w:kern w:val="0"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="en-US" w:eastAsia="%s"/></w:rPr>' % lang)
        # one table per width: a shared grid column would override every row's tcW
        tbls = []
        for w in WIDTHS:
            tbls.append('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
                        '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
                        '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders>'
                        '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                        '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>'
                        '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
                        '<w:p><w:pPr><w:spacing w:line="400" w:lineRule="exact"/><w:jc w:val="center"/>%s</w:pPr>'
                        '<w:r>%s<w:t>%s</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>' % (w, w, w, rpr, rpr, TEXT))
        tbl = "".join(tbls)
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + '>'
                    '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat>'
                    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/>'
                    '</w:compat></w:settings>' % compat)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>' + tbl
               + '<w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" '
               'w:left="1000" w:header="500" w:footer="500" w:gutter="0"/></w:sectPr></w:body></w:document>')
        ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
              '<Default Extension="xml" ContentType="application/xml"/>'
              '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
              '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
              '</Types>')
        rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(os.path.join(OUT, name(lang, compat) + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", rels)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/document.xml", doc)
            z.writestr("word/settings.xml", settings)
    print("wrote", len(ARMS), "docs x", len(WIDTHS), "widths")


def word():
    import win32com.client as w
    res = {}
    for lang, compat in ARMS:
        n = name(lang, compat)
        app = w.DispatchEx("Word.Application")
        app.Visible = False
        try:
            d = app.Documents.Open(os.path.join(OUT, n + ".docx"), ReadOnly=True)
            out = []
            for t in d.Tables:
                r = t.Cell(1, 1).Range
                lines, cur, y = [], "", None
                for i in range(r.Start, r.End - 1):
                    yy = d.Range(i, i).Information(6)
                    if y is not None and yy != y:
                        lines.append(cur); cur = ""
                    y = yy
                    cur += d.Range(i, i + 1).Text
                lines.append(cur)
                out.append([len(x) for x in lines])
            d.Close(False)
        finally:
            app.Quit()
        res[n] = out
        print(n, out, flush=True)
    json.dump(res, open(os.path.join(OUT, "word.json"), "w"), indent=1)


def oxi(exe):
    want = json.load(open(os.path.join(OUT, "word.json")))
    bad = 0
    for lang, compat in ARMS:
        n = name(lang, compat)
        dump = os.path.join(OUT, n + "_oxi.json")
        subprocess.run([exe, os.path.join(OUT, n + ".docx"), dump[:-5], "--dump-layout=" + dump], capture_output=True)
        d = json.load(open(dump, encoding="utf-8"))
        rows = {}
        for pg in d["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and e.get("cell_row_idx") is not None:
                    key = (pg["page"], e["para_idx"])
                    rows.setdefault(key, {}).setdefault(round(e["y"], 1), 0)
                    rows[key][round(e["y"], 1)] += len(e.get("text") or "")
        got = [[v for _, v in sorted(rows[k].items())] for k in sorted(rows)]
        for w, a, b in zip(WIDTHS, want[n], got):
            if a != b:
                bad += 1
                print(n, w, "word", a, "oxi", b)
    print("mismatching rows", bad)


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
