# -*- coding: utf-8 -*-
"""Where does Word put the first baseline of a text box's paragraph under the
section's docGrid -- does a box's line snap to the grid (lines / linesAndChars
/ none), and where does the em box sit in the snapped line?

blind-G ja/policies__20069856 (linesAndChars pitch 360 charSpace 855, pgMar 0,
a page-sized wpg group of text boxes): the 52pt HG創英角ﾎﾟｯﾌﾟ体 title's baseline
is 56.2 below the box's content top (Oxi 52.2), the 36pt ① lines 40.1 (Oxi 36.3),
the 22pt メイリオ lines 34.2 (Oxi 22.3) -- the ① value is (3 cells 54 − 36)/2 +
ascent, i.e. the em box centred in a grid-snapped line, which Oxi applies to
body lines (S1354) but not inside a text box on a linesAndChars grid.

One docx per grid arm: lines360 / lac855 (linesAndChars charSpace=855) /
lacneg (pitch 290, charSpace −2271) / nogrid; each holds six absolutely
positioned text boxes (wrapNone, page-relative, 480x100, default insets), one
paragraph each: メイリオ 22 bold, HG創英角ﾎﾟｯﾌﾟ体 52, HGP創英角ﾎﾟｯﾌﾟ体 36,
ＭＳ 明朝 10.5, ＭＳ 明朝 24, 游ゴシック 18. Read: baseline − box top.

    python tools/metrics/_pb_txbxgrid_gen.py gen
    python tools/metrics/_pb_txbxgrid_gen.py word
    python tools/metrics/_pb_txbxgrid_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_txbxgrid"
GRIDS = {
    "lines360": '<w:docGrid w:type="lines" w:linePitch="360"/>',
    "lac855": '<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="855"/>',
    "lacneg": '<w:docGrid w:type="linesAndChars" w:linePitch="290" w:charSpace="-2271"/>',
    "nogrid": "",
}
# (font, half-points, bold)
FACES = [("メイリオ", 44, True), ("HG創英角ﾎﾟｯﾌﾟ体", 104, False), ("HGP創英角ﾎﾟｯﾌﾟ体", 72, False),
         ("ＭＳ 明朝", 21, False), ("ＭＳ 明朝", 48, False), ("游ゴシック", 36, False)]
BOX_W, BOX_H, BOX_X, BOX_Y0, BOX_STEP = 480.0, 100.0, 36.0, 60.0, 115.0
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="21"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '</w:styles>')


def box(i, font, hp, bold):
    cx, cy = int(BOX_W * 12700), int(BOX_H * 12700)
    px, py = int(BOX_X * 12700), int((BOX_Y0 + i * BOX_STEP) * 12700)
    b = "<w:b/>" if bold else ""
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:hint="eastAsia"/>{b}<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/></w:rPr>'
    para = f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>Ｑ{i}意思表示</w:t></w:r></w:p>'
    return ('<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
            f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{251658240 + i}" behindDoc="0"'
            ' locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{px}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{py}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="{i + 1}" name="TB{i}"/>'
            '<wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:noFill/><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>' + para + '</w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0">'
            '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for name, grid in GRIDS.items():
        body = "<w:p>" + "".join(box(i, f, hp, b) for i, (f, hp, b) in enumerate(FACES)) + "</w:p>"
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134"'
               f' w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>{grid}</w:sectPr></w:body></w:document>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="840"/>'
                    '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:useFELayout/>'
                    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        with zipfile.ZipFile(OUT / f"{name}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(GRIDS), "docs x", len(FACES), "boxes")


def _read(lines):
    res = {}
    for i, (f, hp, b) in enumerate(FACES):
        tx = [l for l in lines if f"Ｑ{i}" in l["t"]]
        if not tx:
            res[f"{f}_{hp / 2:g}"] = None
            continue
        l = tx[0]
        res[f"{f}_{hp / 2:g}"] = {"base_from_top": round(l["y"] - (BOX_Y0 + i * BOX_STEP), 2), "x": round(l["x0"] - BOX_X, 2)}
    return res


def word():
    import win32com.client, line_census as L
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for name in GRIDS:
            t = os.path.join(tempfile.mkdtemp(), f"{name}.docx"); shutil.copy(OUT / f"{name}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{name}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for name in GRIDS:
        res[name] = _read(L.word_lines(OUT / f"{name}.pdf"))
        print(f"{name:9s} word {res[name]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for name in GRIDS:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{name}.docx"))
        wd = word.get(name, {})
        for k, v in o.items():
            wv = wd.get(k)
            flag = "" if v and wv and abs(v["base_from_top"] - wv["base_from_top"]) < 0.6 else "  <-- differs"
            print(f"{name:9s} {k:22s} oxi {v}  word {wv}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
