# -*- coding: utf-8 -*-
"""Does the mandatory empty paragraph AFTER a table that ends a footer (or a
header) take vertical space in Word?

blind-H en/administrative__0067cba8: footer = 1-row 3-cell bordered table
(Calibri 10, pPrDefault after=200 line=276) + one empty Footer paragraph.
Word's PDF puts the row's text baseline at 791.86 on an 841.9pt page with
footer=720tw -- the ROW sits on the footer line, as if the trailing paragraph
were not there; Oxi paints the row 11.2pt higher.

Arms: <where>_<tail>: where = ftr / hdr (the same table + tail paragraph in the
footer or the header); tail = dflt (empty, style spacing after 10 / 1.15),
sz2 (empty, sz 2), after0 (empty, direct spacing after=0 line=240), text
(the paragraph holds «x»), two (two empty paragraphs).  Body: 60 Calibri 11
lines so page 1 is full.  Read, from the PDF: the table text baseline, the
first and last body baselines on page 1.

    python tools/metrics/_pb_ftrtbltail_gen.py gen
    python tools/metrics/_pb_ftrtbltail_gen.py word
    python tools/metrics/_pb_ftrtbltail_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_ftrtbltail"
# ftrbm: the footer again under a 100pt bottom margin, so the margin (not dist + stack) binds the body --
# does the painted row stay dist + stack above the page bottom (0067cba8: bottom 72 > 35.4 + 25.8)?
WHERES = ("ftr", "hdr", "ftrbm")
TAILS = {
    "dflt": '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>',
    "sz2": '<w:p><w:pPr><w:pStyle w:val="Footer"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>',
    "after0": '<w:p><w:pPr><w:pStyle w:val="Footer"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>',
    "text": '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
    "two": '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p><w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>',
    # pre: a text paragraph BEFORE the table (style after 10 / 1.15) -- does its space-after separate it from
    # the table (reference__00c9a41e: Header after=80 then a table; Word's row top sits 4pt below the line)?
    "pre": '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>',
}
LEAD = '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>Lead line before the table</w:t></w:r></w:p>'
ARMS = [f"{w}_{t}" for w in WHERES for t in TAILS]
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" w:pos="9026"/></w:tabs></w:pPr></w:style>'
          '</w:styles>')
CELL = ('<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/>'
        '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>{}</w:t></w:r></w:p></w:tc>')
TABLE = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
         '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>'
         + CELL.format("Date: 7 April 2021") + CELL.format("Version: Final") + CELL.format("Page: 1 of 15") + '</w:tr></w:tbl>')


def part(root, tail):
    lead = LEAD if tail == "pre" else ""
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{root} {T.NS}>{lead}{TABLE}{TAILS[tail]}</w:{root}>'


def doc_xml(where):
    body = "".join(f'<w:p><w:r><w:t>Body line {i + 1} of the filler text that keeps the page full.</w:t></w:r></w:p>' for i in range(60))
    ref = ('<w:headerReference w:type="default" r:id="rIdH"/>' if where == "hdr" else '<w:footerReference w:type="default" r:id="rIdF"/>')
    bottom = 2000 if where == "ftrbm" else 1276
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
            '<w:sectPr>' + ref + f'<w:pgSz w:w="11909" w:h="16834"/><w:pgMar w:top="1440" w:right="1412" w:bottom="{bottom}"'
            ' w:left="1412" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm in ARMS:
        where, tail = arm.split("_")
        name, root, rid, ctype = (("header1", "hdr", "rIdH", "header") if where == "hdr" else ("footer1", "ftr", "rIdF", "footer"))
        ct = T.CT.replace("</Types>", f'<Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{ctype}+xml"/>'
                          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{ctype}" Target="{name}.xml"/>'
                                '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat>'
                    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr(f"word/{name}.xml", part(root, tail))
            z.writestr("word/document.xml", doc_xml(where))
    print("ok", len(ARMS))


def _read(lines):
    tbl = [l["y"] for l in lines if l["t"].startswith("Date")]
    lead = [l["y"] for l in lines if l["t"].startswith("Lead")]
    body = sorted(l["y"] for l in lines if l["t"].startswith("Body"))
    return {"tbl": round(tbl[0], 2) if tbl else None, "lead": round(lead[0], 2) if lead else None,
            "body_first": round(body[0], 2) if body else None,
            "body_last": round(body[-1], 2) if body else None, "n_body": len(body)}


def word():
    import win32com.client, line_census as L
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
        res[arm] = _read(L.word_lines(OUT / f"{arm}.pdf"))
        print(f"{arm:12s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        print(f"{arm:12s} oxi {o}  word {word.get(arm)}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
