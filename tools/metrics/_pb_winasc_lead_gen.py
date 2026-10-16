# -*- coding: utf-8 -*-
"""Where is the baseline of a face whose win ascent EXCEEDS its hhea ascent while
its hhea box is still taller than its win box?

`FontMetrics::baseline_ascent` returns win ascent alone when hhea ascent < win
ascent, dropping the external leading (hhea total - win total) that GDI puts
above the ascent.  Installed Latin faces of that shape: Modern No. 20 (hhea
791/209/70, win 815/209 per 1000: leading 0.046 em) and Haettenschweiler
(1802/246/137, 1884/246 per 2048: 0.027 em).  blind-G en/administrative__0063f992
(Modern No. 20 12pt in a text box): Word 0.54 above Oxi = 0.046 x 12 = 0.55.

Arms <face>_<size>_<where>: face = mod (Modern No. 20) / hae (Haettenschweiler)
/ ari (Arial control); size 12 / 24; where = body (first body line, no grid,
top margin 72) / box (wps box at page 100pt, default insets, no outline).
pPrDefault after 0, line 240 auto, compat 15.  Read: the first baseline minus
the line top (Word PDF; Oxi glyphs).

    python tools/metrics/_pb_winasc_lead_gen.py gen
    python tools/metrics/_pb_winasc_lead_gen.py word
    python tools/metrics/_pb_winasc_lead_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_winasc_lead"
EMU = 12700
BOX_TOP = 100.0
FACES = {"mod": "Modern No. 20", "hae": "Haettenschweiler", "ari": "Arial"}
ARMS = {f"{f}_{s}_{w}": (f, s, w) for f in FACES for s in (12, 24) for w in ("body", "box")}


def rpr(face, size):
    return f'<w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:cs="{face}"/><w:sz w:val="{size * 2}"/><w:szCs w:val="{size * 2}"/></w:rPr>'


def para(face, size, text):
    return f'<w:p><w:pPr>{rpr(face, size)}</w:pPr><w:r>{rpr(face, size)}<w:t>{text}</w:t></w:r></w:p>'


def box(face, size):
    cx, cy = 400 * EMU, 120 * EMU
    return ('<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
            '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(BOX_TOP * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="TB1"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></wp:cNvGraphicFramePr>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
            '<wps:txbx><w:txbxContent>' + para(face, size, "Probe line") + '</w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
            '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>')


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (f, s, w) in ARMS.items():
        if w == "body":
            body = para(FACES[f], s, "Probe line") + para("Arial", 11, "Filler after")
        else:
            body = '<w:p><w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr>' + box(FACES[f], s) + '<w:r><w:t>Anchor para</w:t></w:r></w:p>'
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
               '<w:cols w:space="708"/></w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" '
                    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def top_of(arm):
    return 72.0 if ARMS[arm][2] == "body" else BOX_TOP + 3.6


def word():
    import win32com.client, fitz
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
        pg = fitz.open(str(OUT / f"{arm}.pdf"))[0]; hit = None
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        if c["c"] == "P" and hit is None:
                            hit = (round(c["origin"][1] - top_of(arm), 2), s["font"])
        res[arm] = {"rel": hit[0] if hit else None, "font": hit[1] if hit else None}
        print(f"{arm:14s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]
        hit = [x for x in g if x["char"] == "P"]
        o = {"rel": round(hit[0]["baseline"] - top_of(arm), 2) if hit else None, "font": hit[0]["font_family"] if hit else None}
        flag = "" if o["rel"] is not None and wd[arm]["rel"] is not None and abs(o["rel"] - wd[arm]["rel"]) < 0.2 else "  <-- differs"
        print(f"{arm:14s} oxi {o}  word {wd[arm]}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
