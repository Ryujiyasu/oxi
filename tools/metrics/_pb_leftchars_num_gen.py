# -*- coding: utf-8 -*-
"""Which left indent wins for a numbered paragraph whose style says
leftChars=400/left=840, whose numbering level says left=720/hanging=720, and
which carries a DIRECT <w:ind w:leftChars="0"/>?

blind-G ja/reference__5593c6a7 (text box list «１，体育科がある»…, HGP創英
ﾌﾟﾚｾﾞﾝｽEB 22pt, List Paragraph style leftChars=400 left=840, level ind
left=720 hanging=720, direct ind leftChars=0): Word puts the label at the
text edge and the text 36pt in (the level's hanging), Oxi 83.9pt in (the
style's 400 chars) — «３，２つの体育館と室内プール» wraps and the box grows.

Arms (one paragraph each, same numbering): none (no direct ind), lc0
(leftChars=0), l0 (left=0), lc0l0 (both 0), lc0_nonum (leftChars=0, no
numPr), lc200 (leftChars=200), l360 (left=360).  Read the label x and the
text x from the PDF.

    python tools/metrics/_pb_leftchars_num_gen.py gen
    python tools/metrics/_pb_leftchars_num_gen.py word
    python tools/metrics/_pb_leftchars_num_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_leftchars_num"
ARMS = {
    "none": ("", True), "lc0": ('<w:ind w:leftChars="0"/>', True), "l0": ('<w:ind w:left="0"/>', True),
    "lc0l0": ('<w:ind w:leftChars="0" w:left="0"/>', True), "lc0nonum": ('<w:ind w:leftChars="0"/>', False),
    "lc200": ('<w:ind w:leftChars="200"/>', True), "l360": ('<w:ind w:left="360"/>', True),
    "nonum": ("", False),
    # hangingChars arms: does a chars hanging indent change which twips the numbering contributes?
    "hc100": ('<w:ind w:hangingChars="100"/>', True), "lc200hc100": ('<w:ind w:leftChars="200" w:hangingChars="100"/>', True),
    "hc100nonum": ('<w:ind w:hangingChars="100"/>', False),
}
NUMBERING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + T.NS + '>'
             '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
             '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimalFullWidth"/><w:lvlText w:val="%1，"/><w:lvlJc w:val="left"/>'
             '<w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:lvl></w:abstractNum>'
             '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:eastAsia="MS Gothic" w:cs="MS Gothic"/><w:kern w:val="2"/><w:sz w:val="21"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="a9"><w:name w:val="List Paragraph"/><w:basedOn w:val="a"/><w:pPr><w:ind w:leftChars="400" w:left="840"/></w:pPr></w:style>'
          '</w:styles>')
RPR = '<w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/><w:sz w:val="44"/></w:rPr>'
# LC_BOX=1: the same paragraphs inside a 420pt text box (lIns 7.2pt) -- the real doc's container
BOX = bool(os.environ.get("LC_BOX"))
NAME = "leftchars_box" if BOX else "leftchars"
LEFT = 56.7 + (7.2 if BOX else 0.0)


def boxed(paras):
    cx, cy = int(420 * 12700), int(330 * 12700)
    return ('<w:p><w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
            '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0"'
            ' locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="TB1"/>'
            '<wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:noFill/><a:ln w="6350"><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>' + paras + '</w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0">'
            '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>')


def para(i, ind, num):
    numpr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if num else ""
    return (f'<w:p><w:pPr><w:pStyle w:val="a9"/>{numpr}{ind}<w:jc w:val="left"/>{RPR}</w:pPr>'
            f'<w:r>{RPR}<w:t>Ｑ{i}体育科がある</w:t></w:r></w:p>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    body = "".join(para(i, ind, num) for i, (ind, num) in enumerate(ARMS.values()))
    if BOX:
        body = boxed(body)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134"'
           ' w:left="1134" w:header="720" w:footer="720" w:gutter="0"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
    ct = T.CT.replace("</Types>", '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
                      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
                            '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="840"/>'
                '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    with zipfile.ZipFile(OUT / f"{NAME}.docx", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING); z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def _read(lines, left=LEFT):
    res = {}
    for i, arm in enumerate(ARMS):
        tx = [l for l in lines if f"Ｑ{i}" in l["t"]]
        if not tx:
            res[arm] = None
            continue
        l = tx[0]
        # the label («１，») is a separate line object in the PDF; the text line starts at Ｑ
        txt_x = l["x0"] + (l["t"].find(f"Ｑ{i}") / max(len(l["t"]), 1)) * (l["x1"] - l["x0"]) if l["t"].find(f"Ｑ{i}") > 0 else l["x0"]
        labels = [m for m in lines if abs(m["y"] - l["y"]) < 0.5 and m is not l]
        res[arm] = {"text_x": round(txt_x - left, 1), "label_x": round(min(m["x0"] for m in labels) - left, 1) if labels else None}
    return res


def word():
    import win32com.client, line_census as L
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        t = os.path.join(tempfile.mkdtemp(), f"{NAME}.docx"); shutil.copy(OUT / f"{NAME}.docx", t)
        d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
        d.SaveAs2(str((OUT / f"{NAME}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = _read(L.word_lines(OUT / f"{NAME}.pdf"))
    for arm in ARMS:
        print(f"{arm:9s} word {res[arm]}", flush=True)
    (OUT / f"word_{NAME}.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / f"word_{NAME}.json").read_text(encoding="utf-8"))
    res = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{NAME}.docx"))
    for arm in ARMS:
        print(f"{arm:9s} oxi {res[arm]}   word {word.get(arm)}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
