# -*- coding: utf-8 -*-
"""A text box whose bodyPr says wrap="none" (no wrapping inside the shape) with a
fixed extent wider than its text: where does Word put the text horizontally?

blind-G ja/reference__5593c6a7: «部活が盛んな柴田高校» (HGS明朝E 36, no jc, box
629.2 wide, lIns 5.85, wrap="none", anchorCtr="0"): Word's text starts 134pt in
(x 183.6 in a box at 49.7) = centred in the box; Oxi starts at the left inset.

Arms (one 480x60 box each, page-relative at x=36, ＭＳ 明朝 24pt «Ｑn柴田高校»
~144pt wide): wrapnone_left (no jc), wrapnone_center, wrapnone_right,
wrapnone_long (a 26-char text wider than the box), square_left (wrap="square"
control), wrapnone_ctr1 (anchorCtr="1"). Read: text x0 − box x.

    python tools/metrics/_pb_wrapnone_gen.py gen
    python tools/metrics/_pb_wrapnone_gen.py word
    python tools/metrics/_pb_wrapnone_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_wrapnone"
# (wrap, jc, long text, anchorCtr)
ARMS = {
    "wrapnone_left": ("none", None, False, "0"), "wrapnone_center": ("none", "center", False, "0"),
    "wrapnone_right": ("none", "right", False, "0"), "wrapnone_long": ("none", None, True, "0"),
    "square_left": ("square", None, False, "0"), "wrapnone_ctr1": ("none", None, False, "1"),
    # H-aligned boxes (relativeFrom margin, align center / right): does a wrap="none"
    # box take its TEXT's width for the alignment (the 5593c6a7 reading)?
    "wrapnone_hcenter": ("none", None, False, "0", "center"), "square_hcenter": ("square", None, False, "0", "center"),
    "wrapnone_hright": ("none", None, False, "0", "right"), "wrapnone_hcenter_jcc": ("none", "center", False, "0", "center"),
}
BOX_W, BOX_H, BOX_X, BOX_Y0, BOX_STEP = 480.0, 60.0, 36.0, 60.0, 80.0
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="21"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '</w:styles>')


def box(i, wrap, jc, long, ctr, halign=None):
    cx, cy = int(BOX_W * 12700), int(BOX_H * 12700)
    px, py = int(BOX_X * 12700), int((BOX_Y0 + i * BOX_STEP) * 12700)
    pos_h = (f'<wp:positionH relativeFrom="margin"><wp:align>{halign}</wp:align></wp:positionH>' if halign
             else f'<wp:positionH relativeFrom="page"><wp:posOffset>{px}</wp:posOffset></wp:positionH>')
    text = f"Ｑ{i}柴田高校" + ("は運動部も学芸部も部活動がとても盛んです" if long else "")
    jcx = f'<w:jc w:val="{jc}"/>' if jc else ""
    rpr = '<w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr>'
    para = f'<w:p><w:pPr>{jcx}{rpr}</w:pPr><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>'
    return ('<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
            f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{251658240 + i}" behindDoc="0"'
            ' locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            + pos_h +
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{py}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="{i + 1}" name="TB{i}"/>'
            '<wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:noFill/><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>' + para + '</w:txbxContent></wps:txbx>'
            f'<wps:bodyPr rot="0" vert="horz" wrap="{wrap}" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="{ctr}">'
            '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    body = "<w:p>" + "".join(box(i, *a) for i, a in enumerate(ARMS.values())) + "</w:p>"
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134"'
           ' w:left="1134" w:header="720" w:footer="720" w:gutter="0"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="840"/>'
                '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:useFELayout/>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    with zipfile.ZipFile(OUT / "wrapnone.docx", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def _read(lines):
    res = {}
    for i, arm in enumerate(ARMS):
        tx = [l for l in lines if f"Ｑ{i}" in l["t"]]
        off = 0.0 if len(ARMS[arm]) > 4 else BOX_X
        res[arm] = {"x0": round(tx[0]["x0"] - off, 2), "x1": round(tx[0]["x1"] - off, 2), "y": round(tx[0]["y"] - (BOX_Y0 + i * BOX_STEP), 2), "n": len(tx)} if tx else None
    return res


def word():
    import win32com.client, line_census as L
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        t = os.path.join(tempfile.mkdtemp(), "wrapnone.docx"); shutil.copy(OUT / "wrapnone.docx", t)
        d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
        d.SaveAs2(str((OUT / "wrapnone.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = _read(L.word_lines(OUT / "wrapnone.pdf"))
    for arm in ARMS:
        print(f"{arm:16s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    res = _read(L.oxi_lines(os.path.abspath(exe), OUT / "wrapnone.docx"))
    for arm in ARMS:
        o, wd = res[arm], word.get(arm)
        flag = "" if o and wd and abs(o["x0"] - wd["x0"]) < 1.0 else "  <-- differs"
        print(f"{arm:16s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
