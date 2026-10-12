# -*- coding: utf-8 -*-
"""Where in a multiple-spaced Latin table-cell line does Word put the glyphs:
is the extra leading above or below the text?

`_pb_ftrtbltail_gen.py` hdr_* arms: a bordered Calibri-10 header row (style
line 276) has its baseline 9.96 below the row's top border in Word (45.96 −
36.0) and 11.26 in Oxi — the 1.3 is roughly the extra leading's share Oxi puts
above the text. The footer rows of the same probe match within 0.6.

Arms: <where>_<line>_<font>: where = body / hdr (the table in the body or the
header); line = 240 / 276 / 360 / 480; font = cal10 (Calibri 10) / tnr12
(Times New Roman 12).  One bordered 1-row table, cell text «Date: 7 April
2021»; read the baseline and the row's top border from the PDF.

    python tools/metrics/_pb_cellmult_top_gen.py gen
    python tools/metrics/_pb_cellmult_top_gen.py word
    python tools/metrics/_pb_cellmult_top_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_cellmult_top"
FONTS = {"cal10": ("Calibri", 20), "tnr12": ("Times New Roman", 24)}
LINES = ("240", "276", "360", "480")
WHERES = ("body", "hdr")
ARMS = [f"{w}_{l}_{f}" for w in WHERES for l in LINES for f in FONTS]
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>')


def table(line, font):
    fam, sz = FONTS[font]
    rpr = f'<w:rPr><w:rFonts w:ascii="{fam}" w:hAnsi="{fam}"/><w:sz w:val="{sz}"/></w:rPr>'
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr><w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/>{rpr}</w:pPr><w:r>{rpr}<w:t>Date: 7 April 2021</w:t></w:r></w:p>'
            '</w:tc></w:tr></w:tbl>')


def doc_xml(where, line, font):
    body = "".join(f'<w:p><w:r><w:t>Body line {i + 1}.</w:t></w:r></w:p>' for i in range(3))
    if where == "body":
        body = table(line, font) + body
    ref = '<w:headerReference w:type="default" r:id="rIdH"/>' if where == "hdr" else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
            '<w:sectPr>' + ref + '<w:pgSz w:w="11909" w:h="16834"/><w:pgMar w:top="1440" w:right="1412" w:bottom="1276"'
            ' w:left="1412" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm in ARMS:
        where, line, font = arm.split("_")
        ct, drels, parts = T.CT, T.DRELS, {}
        if where == "hdr":
            ct = ct.replace("</Types>", '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')
            drels = drels.replace("</Relationships>", '<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>')
            parts["word/header1.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ' + T.NS + '>' + table(line, font) + '<w:p/></w:hdr>'
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            for n, c in parts.items():
                z.writestr(n, c)
            z.writestr("word/document.xml", doc_xml(where, line, font))
    print("ok", len(ARMS))


def _word_read(pdf):
    import fitz
    pg = fitz.open(pdf)[0]
    base = next((l["spans"][0]["origin"][1] for b in pg.get_text("dict")["blocks"] for l in b.get("lines", [])
                 if "".join(s["text"] for s in l["spans"]).startswith("Date")), None)
    tops = [d["rect"].y0 for d in pg.get_drawings() if d["rect"].width > 50 and d["rect"].height < 1.5]
    bots = sorted(set(round(t, 2) for t in tops))
    return {"base_from_top": round(base - bots[0], 2) if base and bots else None, "row_h": round(bots[1] - bots[0], 2) if len(bots) > 1 else None}


def _oxi_read(exe, docx):
    import subprocess, tempfile
    t = tempfile.mkdtemp(); out = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-layout=" + out], capture_output=True)
    els = json.load(open(out, encoding="utf-8"))["pages"][0]["elements"]
    tops = sorted(set(round(e["y"], 2) for e in els if e["type"] == "border" and e.get("w", 0) > 50))
    import line_census as L
    base = next((l["y"] for l in L.oxi_lines(exe, docx) if l["t"].startswith("Date")), None)
    return {"base_from_top": round(base - tops[0], 2) if base is not None and tops else None, "row_h": round(tops[1] - tops[0], 2) if len(tops) > 1 else None}


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in ARMS:
        res[arm] = _word_read(OUT / f"{arm}.pdf")
        print(f"{arm:16s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _oxi_read(os.path.abspath(exe), OUT / f"{arm}.docx")
        wd = word.get(arm) or {}
        flag = "" if o["base_from_top"] is not None and wd.get("base_from_top") is not None and abs(o["base_from_top"] - wd["base_from_top"]) < 0.4 else "  <-- differs"
        print(f"{arm:16s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
