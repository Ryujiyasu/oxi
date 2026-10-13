# -*- coding: utf-8 -*-
"""Line pitch inside a compatLnSpc="1" text box for CJK faces whose usWin box
is not 1 em (游明朝, メイリオ) and for a face under w14:shadow / w:outline.

blind-G ja/reference__5593c6a7 (no-type docGrid linePitch 360, every box
compatLnSpc=1): the «令和2年度 / 柴田高校 夢実現» box (游明朝 14 / 13 via the
theme) has Word pitches 22.2 / 21.7 (= win 1.287 × 83/64 × fs) where Oxi's
CLNSP box law gives 18.2 / 16.8 (the ＭＳ-calibrated fs + 0.25 / 0.27 fs
constants); the numbered «１，体育科がある» box (HGP創英ﾌﾟﾚｾﾞﾝｽEB 22, bold,
outline, w14:shadow) has Word pitch 31.15 (= 22 × 1.416) where 22 × 83/64 is
28.5 — the shadow / outline growth?

One 480x170 wrapNone box per arm, page-positioned, compatLnSpc="1", three
identical paragraphs; read the baseline pitch (and the first baseline from the
box top). Arms: yumin13, yumin14, msmin10.5, meiryo12, hgp22, hgp22_shadow
(w14:shadow blurRad 38100 dist 22860 dir 5400000 + w:outline + w:b),
hgp22_outline (w:outline + w:b only), hgp22_bold (w:b only).

    python tools/metrics/_pb_clnsp_cjk_gen.py gen
    python tools/metrics/_pb_clnsp_cjk_gen.py word
    python tools/metrics/_pb_clnsp_cjk_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_clnsp_cjk"
SHADOW = ('<w14:shadow w14:blurRad="38100" w14:dist="22860" w14:dir="5400000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="b">'
          '<w14:srgbClr w14:val="000000"><w14:alpha w14:val="63000"/></w14:srgbClr></w14:shadow>')
# (font, half-points, extra rPr xml)
ARMS = {
    "yumin13": ("游明朝", 26, ""), "yumin14": ("游明朝", 28, ""), "msmin10.5": ("ＭＳ 明朝", 21, ""), "meiryo12": ("メイリオ", 24, ""),
    "hgp22": ("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 44, ""), "hgp22_shadow": ("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 44, "<w:b/><w:outline/>" + SHADOW),
    "hgp22_outline": ("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 44, "<w:b/><w:outline/>"), "hgp22_bold": ("HGP創英ﾌﾟﾚｾﾞﾝｽEB", 44, "<w:b/>"),
}
BOX_W, BOX_H, BOX_X, BOX_Y0, BOX_STEP = 480.0, 170.0, 36.0, 60.0, 185.0
NS14 = ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:kern w:val="2"/><w:sz w:val="21"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '</w:styles>')


def box(i, font, hp, extra):
    cx, cy = int(BOX_W * 12700), int(BOX_H * 12700)
    px, py = int(BOX_X * 12700), int((BOX_Y0 + i * BOX_STEP) * 12700)
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:hint="eastAsia"/>{extra}<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/></w:rPr>'
    paras = "".join(f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>Ｑ{i}{k}柴田高校夢実現</w:t></w:r></w:p>' for k in range(3))
    return ('<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
            f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{251658240 + i}" behindDoc="0"'
            ' locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{px}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{py}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="{i + 1}" name="TB{i}"/>'
            '<wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:noFill/><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>' + paras + '</w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0" compatLnSpc="1">'
            '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    body = "<w:p>" + "".join(box(i, f, hp, ex) for i, (f, hp, ex) in enumerate(ARMS.values())) + "</w:p>"
    ns = T.NS if "w14=" in T.NS else T.NS + NS14
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + ns + "><w:body>" + body +
           '<w:sectPr><w:pgSz w:w="14572" w:h="20639"/><w:pgMar w:top="851" w:right="851" w:bottom="851"'
           ' w:left="851" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                      '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                            '<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="840"/>'
                '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:useFELayout/>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    theme = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">'
             '<a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
             '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
             '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
             '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
             '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="游ゴシック Light"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/></a:majorFont>'
             '<a:minorFont><a:latin typeface="游明朝"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游明朝"/></a:minorFont></a:fontScheme>'
             '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
             '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
             '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
             '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
             '</a:themeElements></a:theme>')
    with zipfile.ZipFile(OUT / "clnspcjk.docx", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", settings); z.writestr("word/theme/theme1.xml", theme); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def _read(lines):
    res = {}
    for i, arm in enumerate(ARMS):
        ls = sorted([l for l in lines if f"Ｑ{i}" in l["t"]], key=lambda l: l["y"])
        top = BOX_Y0 + i * BOX_STEP
        res[arm] = {"first": round(ls[0]["y"] - top, 2), "pitch": [round(ls[k + 1]["y"] - ls[k]["y"], 2) for k in range(len(ls) - 1)]} if ls else None
    return res


def word():
    import win32com.client, line_census as L
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        t = os.path.join(tempfile.mkdtemp(), "clnspcjk.docx"); shutil.copy(OUT / "clnspcjk.docx", t)
        d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
        d.SaveAs2(str((OUT / "clnspcjk.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = _read(L.word_lines(OUT / "clnspcjk.pdf"))
    for arm in ARMS:
        print(f"{arm:14s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    res = _read(L.oxi_lines(os.path.abspath(exe), OUT / "clnspcjk.docx"))
    for arm in ARMS:
        o, wd = res[arm], word.get(arm)
        flag = "" if o and wd and o["pitch"] and wd["pitch"] and abs(o["pitch"][0] - wd["pitch"][0]) < 0.4 and abs(o["first"] - wd["first"]) < 0.6 else "  <-- differs"
        print(f"{arm:14s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
