# -*- coding: utf-8 -*-
"""Where does Word put the FIRST baseline of a CJK-face text box?

blind-G ja/administrative__1a08ddac (ＭＳ ゴシック 14 bold title box, 2pt line,
anchor t) sits 1.06 high in Oxi; ja/reference__5593c6a7 (HGP創英角ﾎﾟｯﾌﾟ体 20-22pt
boxes, anchor t) 1.52-1.58 high on five lines; ja/policies__20069856 (HG創英角
ﾎﾟｯﾌﾟ体 52pt, anchor ctr) 1.53 high -- a constant, not a size ratio.

One floating wps text box per arm, positionV relativeFrom="page" posOffset
= 100pt (so the box top is 100), extent 400 x 120pt, default insets
(91440 / 45720), wrapNone, anchored in the first body paragraph.  Arms
<face>_<size>_<anchor>[_ln2][_nogrid]: face = msgo (ＭＳ ゴシック) / pop
(HGP創英角ﾎﾟｯﾌﾟ体) / msmin (ＭＳ 明朝); size = 14 / 22 / 52 pt; anchor t or
ctr; ln2 = 2pt outline; nogrid = no docGrid (default: lines 360).  Read: the
text baseline minus the box top (Word PDF; Oxi glyph dump).

    python tools/metrics/_pb_txbx_cjk_base_gen.py gen
    python tools/metrics/_pb_txbx_cjk_base_gen.py word
    python tools/metrics/_pb_txbx_cjk_base_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_txbx_cjk_base"
EMU = 12700
FACES = {"msgo": "ＭＳ ゴシック", "pop": "HGP創英角ﾎﾟｯﾌﾟ体", "msmin": "ＭＳ 明朝"}
TOP_PT = 100
# (face, size, anchor, ln2, grid)
ARMS = {}
for f in FACES:
    for s in (14, 22, 52):
        ARMS[f"{f}_{s}_t"] = (f, s, "t", False, True)
ARMS["pop_52_ctr"] = ("pop", 52, "ctr", False, True)
ARMS["msgo_14_t_ln2"] = ("msgo", 14, "t", 25400, True)
ARMS["msgo_14_t_nogrid"] = ("msgo", 14, "t", False, False)
ARMS["pop_22_t_nogrid"] = ("pop", 22, "t", False, False)
# outline width sweep (EMU): 9525 = 0.75pt (the default), 25400 = 2pt, 50800 = 4pt; a bare <a:ln> with
# a fill and no w (Word's default width); anchor ctr and no-grid with the 2pt line
ARMS["msgo_14_t_ln075"] = ("msgo", 14, "t", 9525, True)
ARMS["msgo_14_t_ln4"] = ("msgo", 14, "t", 50800, True)
ARMS["msgo_14_t_lnbare"] = ("msgo", 14, "t", -1, True)
ARMS["msgo_14_t_ln2_nogrid"] = ("msgo", 14, "t", 25400, False)
ARMS["pop_52_ctr_ln2"] = ("pop", 52, "ctr", 25400, True)
ARMS["pop_22_t_ln2"] = ("pop", 22, "t", 25400, True)
# a declared width with NO fill (20069856's «要 不要» box: <a:ln w="6350"><a:noFill/>)
ARMS["msgo_14_t_ln2_nofill"] = ("msgo", 14, "t", -2, True)


def box(face, size, anchor, ln2):
    cx, cy = 400 * EMU, 120 * EMU
    if ln2 == -2:
        ln = '<a:ln w="25400"><a:noFill/></a:ln>'
    elif ln2 == -1:
        ln = '<a:ln><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>'
    elif ln2:
        ln = f'<a:ln w="{ln2}"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>'
    else:
        ln = '<a:ln><a:noFill/></a:ln>'
    rpr = f'<w:rPr><w:rFonts w:ascii="{face}" w:eastAsia="{face}" w:hAnsi="{face}" w:hint="eastAsia"/><w:b/><w:sz w:val="{size * 2}"/></w:rPr>'
    return ('<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
            '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{TOP_PT * EMU}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="TB1"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></wp:cNvGraphicFramePr>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp><wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr><wps:spPr bwMode="auto">'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' + ln + '</wps:spPr>'
            '<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/>' + rpr + '</w:pPr><w:r>' + rpr + '<w:t>様式例通知書</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
            f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="{anchor}" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
            '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>')


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="21"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (f, s, an, ln2, grid) in ARMS.items():
        body = '<w:p><w:pPr><w:jc w:val="right"/></w:pPr>' + box(FACES[f], s, an, ln2) + '</w:p>'
        body += "".join(f'<w:p><w:r><w:t>本文{i}行目</w:t></w:r></w:p>' for i in range(1, 6))
        g = '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ''
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1701" w:right="1701" w:bottom="1418" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>'
               '<w:cols w:space="720"/>' + g + '</w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" '
                    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


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
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if "様式" in s["text"]:
                        hit = (round(s["origin"][1], 2), s["font"], round(s["size"], 2))
        res[arm] = {"base": hit[0] if hit else None, "rel": round(hit[0] - TOP_PT, 2) if hit else None, "font": hit[1] if hit else None}
        print(f"{arm:18s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]
        hit = [x for x in g if x["char"] == "様"]
        base = round(hit[0]["baseline"], 2) if hit else None
        o = {"base": base, "rel": round(base - TOP_PT, 2) if base else None, "font": hit[0]["font_family"] if hit else None}
        flag = "" if o["rel"] is not None and wd[arm]["rel"] is not None and abs(o["rel"] - wd[arm]["rel"]) < 0.3 else "  <-- differs"
        print(f"{arm:18s} oxi {o}  word {wd[arm]}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
