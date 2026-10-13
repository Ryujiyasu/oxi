# -*- coding: utf-8 -*-
"""How narrow a lane beside a wrapSquare text box will Word still flow body
text into?

blind-G ja/correspondence__14012ac2 (column 489pt; 457.5pt wrapSquare boxes at
margin+16.1 / page+73 with distL/R 9): the lanes beside them are 6.4 / 7.3pt
wide and Word flows nothing into them (the body resumes below each box); Oxi
flows single characters into the lanes («そⅣの参他加…»).

One docx per lane width L in {4, 8, 12, 16, 20, 24, 30, 40} (pt): a 60pt-tall
wrapSquare box anchored to paragraph 1 (posV paragraph 0, posH margin 0,
distL/distR 9), width = column − 9 − L, followed by 6 body lines of ＭＳ 明朝
10.5 CJK text. Read: do any body glyphs sit right of the box while vertically
inside it (the lane is used), and where the first body line after the box
starts.

    python tools/metrics/_pb_lanewidth_gen.py gen
    python tools/metrics/_pb_lanewidth_gen.py word
    python tools/metrics/_pb_lanewidth_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_lanewidth"
LANES = [4, 8, 12, 16, 17, 18, 19, 20, 24, 30, 40, "far_both", "far_right", "far_left"]
MARGIN_L, MARGIN_R, PAGE_W = 56.7, 56.7, 595.3
COLUMN = PAGE_W - MARGIN_L - MARGIN_R
BOX_H, DIST = 60.0, 9.0
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="21"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '</w:styles>')
BODY = "愛知県中学校総合体育大会について二月に行いました評議員会にて決定しました内容についてお伝えします。"


FAR = {"far_both": ("page", 73.0, COLUMN - (73.0 - MARGIN_L) - DIST - 6.2, 120.0),   # lanes 7.3 / 6.2
       "far_right": ("margin", 0.0, COLUMN - DIST - 10.0, 120.0),                   # lanes 0 / 10
       "far_left": ("page", 73.0, COLUMN - (73.0 - MARGIN_L) - DIST - 40.0, 120.0)}  # lanes 7.3 / 40


def box(lane):
    if lane in FAR:
        rel, hx, w, vy = FAR[lane]
    else:
        rel, hx, w, vy = "margin", 0.0, COLUMN - DIST - lane, 0.0
    cx, cy = int(w * 12700), int(BOX_H * 12700)
    d = int(DIST * 12700)
    para = '<w:p><w:r><w:t>箱の中の文</w:t></w:r></w:p>'
    return ('<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
            f'<wp:anchor distT="0" distB="0" distL="{d}" distR="{d}" simplePos="0" relativeHeight="251658240" behindDoc="0"'
            ' locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="{rel}"><wp:posOffset>{int(hx * 12700)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{int(vy * 12700)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="TB1"/>'
            '<wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:noFill/><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>' + para + '</w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0">'
            '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="840"/>'
                '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:useFELayout/>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    for lane in LANES:
        body = "<w:p>" + box(lane) + f"<w:r><w:t>Ａ{BODY}</w:t></w:r></w:p>"
        body += "".join(f"<w:p><w:r><w:t>{chr(ord('Ｂ') + i)}{BODY}</w:t></w:r></w:p>" for i in range(5 if lane not in FAR else 12))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134"'
               ' w:left="1134" w:header="720" w:footer="720" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"lane{lane}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(LANES))


def _read(lines, lane):
    if lane in FAR:
        rel, hx, w, vy = FAR[lane]
        box_left = (hx if rel == "page" else MARGIN_L + hx); box_right = box_left + w
        box_top = 56.7 + vy; box_bot = box_top + BOX_H
        body = sorted([l for l in lines if "箱" not in l["t"]], key=lambda l: l["y"])
        inside = [l for l in body if box_top - 2 < l["y"] < box_bot + 12]
        beside = [l for l in inside if l["x1"] < box_left + 1.0 or l["x0"] > box_right - 1.0]
        over = [l for l in inside if not (l["x1"] < box_left + 1.0 or l["x0"] > box_right - 1.0)]
        after = [l for l in body if l["y"] > box_bot + 12]
        return {"beside": len(beside), "beside_chars": "".join(l["t"] for l in beside)[:10], "over": len(over),
                "resume_y": round(after[0]["y"], 2) if after else None, "n_body": len(body)}
    box_right = MARGIN_L + COLUMN - DIST - lane
    box_top, box_bot = 56.7, 56.7 + BOX_H
    in_lane = [l for l in lines if l["x0"] > box_right - 1.0 and box_top - 2 < l["y"] < box_bot + 12 and "箱" not in l["t"]]
    body = sorted([l for l in lines if "箱" not in l["t"]], key=lambda l: l["y"])
    first = body[0] if body else None
    return {"lane_used": len(in_lane), "lane_chars": "".join(l["t"] for l in in_lane)[:10], "first_body_y": round(first["y"], 2) if first else None,
            "first_body_x": round(first["x0"], 1) if first else None, "n_body": len(body)}


def word():
    import win32com.client, line_census as L
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for lane in LANES:
            t = os.path.join(tempfile.mkdtemp(), f"lane{lane}.docx"); shutil.copy(OUT / f"lane{lane}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"lane{lane}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for lane in LANES:
        res[str(lane)] = _read(L.word_lines(OUT / f"lane{lane}.pdf"), lane)
        print(f"lane {lane!s:9} word {res[str(lane)]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for lane in LANES:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"lane{lane}.docx"), lane)
        wd = word.get(str(lane))
        flag = "" if wd and all(o.get(k) == wd.get(k) for k in ("lane_used", "beside", "over")) and abs((o.get("first_body_y") or o.get("resume_y") or 0) - (wd.get("first_body_y") or wd.get("resume_y") or 0)) < 1.0 else "  <-- differs"
        print(f"lane {lane!s:9} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
