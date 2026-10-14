# -*- coding: utf-8 -*-
"""Does a FLOATING (anchored) picture between two header paragraphs break the
first paragraph's space-after?  blind-G en/policies__009ad56e: first-page
header = «Commonwealth of Massachusetts» (Times New Roman 18 bold, direct
after=40) + a page-anchored seal + «Disabled Persons Protection Commission»
(Times 16).  Word's pitch is 20.76 (= 18pt descent + 16pt ascent + 2pt after);
Oxi's 18.84 -- the IR puts the anchored picture between the paragraphs as its
own Image block and the paint drops the previous paragraph's after there
(the reservation keeps it).

Arms a<after>_<anchor>: after = 0 / 40 / 200 twips (direct on paragraph 1);
anchor = none / p1 (the picture anchored in paragraph 1) / p2 (in paragraph
2).  Letter, margins 1440 / bottom 851, header 720, docGrid linePitch 360
(no type), compat 14, docDefaults Calibri 11 with an EMPTY pPrDefault, Header
style = tabs only (the document's shape).  Read: the two header baselines and
the first body baseline (Word PDF; Oxi glyph dump).

    python tools/metrics/_pb_hdrafter_anchor_gen.py gen
    python tools/metrics/_pb_hdrafter_anchor_gen.py word
    python tools/metrics/_pb_hdrafter_anchor_gen.py oxi <renderer.exe>
"""
import json, os, shutil, struct, subprocess, sys, tempfile, zipfile, zlib
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402
import line_census as L  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_hdrafter_anchor"
NSW = " ".join(f'xmlns:{p}="{u}"' for p, u in (("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
                                              ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
                                              ("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
                                              ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")) if f"xmlns:{p}=" not in T.NS)
TNR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
ARMS = [f"a{a}_{an}" for a in (0, 40, 200) for an in ("none", "p1", "p2")]
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs>'
          '<w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs></w:pPr></w:style></w:styles>')


def png(w=96, h=96):
    raw = b"".join(b"\x00" + bytes([60, 60, 140] * w) for _ in range(h))
    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xffffffff)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def anchor_run():
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="301" name="Seal"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            '<pic:nvPicPr><pic:cNvPr id="301" name="Seal"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def header(after, anchor):
    p1 = ('<w:p><w:pPr><w:pStyle w:val="Header"/><w:tabs><w:tab w:val="clear" w:pos="4680"/></w:tabs>' + (f'<w:spacing w:after="{after}"/>' if after else '') +
          '<w:ind w:left="360"/><w:jc w:val="center"/></w:pPr>' + (anchor_run() if anchor == "p1" else '') +
          '<w:r><w:rPr>' + TNR + '<w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr><w:t>Commonwealth of Massachusetts</w:t></w:r></w:p>')
    p2 = ('<w:p><w:pPr><w:pStyle w:val="Header"/><w:tabs><w:tab w:val="clear" w:pos="4680"/></w:tabs><w:ind w:left="360"/><w:jc w:val="center"/></w:pPr>' +
          (anchor_run() if anchor == "p2" else '') +
          '<w:r><w:rPr>' + TNR + '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr><w:t>Disabled Persons Protection Commission</w:t></w:r></w:p>')
    p3 = '<w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr></w:p>'
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ' + T.NS + " " + NSW + ">" + p1 + p2 + p3 + "</w:hdr>"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    body = "".join(f'<w:p><w:r><w:t>Body line {i + 1} of the policy text that fills the page.</w:t></w:r></w:p>' for i in range(20))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + " " + NSW + "><w:body>" + body +
           '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="851"'
           ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>')
    for arm in ARMS:
        a, an = arm.split("_"); after = int(a[1:])
        ct = T.CT.replace("</Types>", '<Default Extension="png" ContentType="image/png"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
                          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
                                '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        hrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/header1.xml", header(after, an))
            z.writestr("word/_rels/header1.xml.rels", hrels); z.writestr("word/media/image1.png", png()); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def _read(lines):
    h = sorted(l["y"] for l in lines if l["t"].startswith(("Commonwealth", "Disabled")))
    body = sorted(l["y"] for l in lines if l["t"].startswith("Body"))
    return {"hdr": [round(y, 2) for y in h], "pitch": round(h[1] - h[0], 2) if len(h) == 2 else None, "body_first": round(body[0], 2) if body else None}


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
    res = {arm: _read(L.word_lines(OUT / f"{arm}.pdf")) for arm in ARMS}
    for arm, v in res.items():
        print(f"{arm:10s} word {v}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        flag = "" if o["pitch"] is not None and wd[arm]["pitch"] is not None and abs(o["pitch"] - wd[arm]["pitch"]) < 0.3 else "  <-- differs"
        print(f"{arm:10s} oxi {o}  word {wd[arm]}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
