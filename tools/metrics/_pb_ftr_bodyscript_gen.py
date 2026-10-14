# -*- coding: utf-8 -*-
"""Is a footer's placement decided by the BODY's script or by the footer's own
faces?  Oxi gates its Latin footer stack (S806 / S886 / S992 / the pBdr paint
room) on `doc_body_has_real_cjk`; blind-G ja/reports__597cb86c (Calibri 12
footer with a pBdr top, 443 CJK characters in the body) sits 1.01 low and
ja/forms__02157cbe (Century 10.5 page number + an empty paragraph) 1.13 high,
and both land on Word the moment the body is replaced by Latin filler.

Arms <body>_<ftr>: body = lat (Calibri filler) / cjk (the same filler plus a
ＭＳ 明朝 paragraph «日本語の段落です。»); ftr = calbdr (Calibri 12, pBdr top
sz 4, no pStyle, no pPrDefault -- the report) / cent2 (Century 10.5 «1» centred
+ an empty paragraph, ＭＳ 明朝 eastAsia, snapToGrid 0 -- the form) / msmin
(ＭＳ 明朝 10.5 «脚注の行» one paragraph).  A4, margins 1440, footer 720,
compat 15.  Read: the footer baseline (Word PDF; Oxi glyph dump).

    python tools/metrics/_pb_ftr_bodyscript_gen.py gen
    python tools/metrics/_pb_ftr_bodyscript_gen.py word
    python tools/metrics/_pb_ftr_bodyscript_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_ftr_bodyscript"
MS = "ＭＳ 明朝"
BODIES = ("lat", "cjk")
FTRS = ("calbdr", "cent2", "msmin")
ARMS = [f"{b}_{f}" for b in BODIES for f in FTRS]


def styles(ftr):
    if ftr == "calbdr":
        rf = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/>'
        ppd = ""
    else:
        rf = '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="' + MS + '" w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:kern w:val="2"/>'
        ppd = "<w:pPrDefault/>"
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>' + rf +
            '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>' + ppd + '</w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="Ftr"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4252"/>'
            '<w:tab w:val="right" w:pos="8504"/></w:tabs><w:snapToGrid w:val="0"/></w:pPr></w:style></w:styles>')


def footer(ftr):
    if ftr == "calbdr":
        body = ('<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4"/></w:pBdr></w:pPr><w:r><w:rPr/><w:t xml:space="preserve">Footer created by </w:t></w:r>'
                '<w:r><w:rPr><w:color w:val="2980b9"/><w:u w:val="single"/></w:rPr><w:t>FullPicture.app</w:t></w:r></w:p>')
    elif ftr == "cent2":
        body = ('<w:p><w:pPr><w:pStyle w:val="Ftr"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r></w:p>'
                '<w:p><w:pPr><w:pStyle w:val="Ftr"/></w:pPr></w:p>')
    else:
        body = '<w:p><w:pPr><w:pStyle w:val="Ftr"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>脚注の行</w:t></w:r></w:p>'
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + T.NS + '>' + body + '</w:ftr>'


def doc_xml(body, ftr):
    paras = "".join('<w:p><w:r><w:t>Body line %d of the filler text that keeps the page.</w:t></w:r></w:p>' % (i + 1) for i in range(20))
    if body == "cjk":
        paras += '<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>日本語の段落です。</w:t></w:r></w:p>'
    grid = '<w:docGrid w:type="linesAndChars" w:linePitch="373"/>' if ftr == "cent2" else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + paras +
            '<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>' + grid + '</w:sectPr></w:body></w:document>')


SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm in ARMS:
        b, f = arm.split("_")
        ct = T.CT.replace("</Types>", '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
                                '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles(f))
            z.writestr("word/settings.xml", SETTINGS); z.writestr("word/footer1.xml", footer(f)); z.writestr("word/document.xml", doc_xml(b, f))
    print("ok", len(ARMS))


def _read(lines):
    f = sorted(l["y"] for l in lines if l["y"] > 740)
    body = sorted(l["y"] for l in lines if l["t"].startswith("Body"))
    return {"ftr": [round(y, 2) for y in f], "body_first": round(body[0], 2) if body else None}


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
        wd = word.get(arm) or {}
        flag = "" if o["ftr"] and wd.get("ftr") and abs(o["ftr"][0] - wd["ftr"][0]) < 0.6 else "  <-- differs"
        print(f"{arm:12s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
