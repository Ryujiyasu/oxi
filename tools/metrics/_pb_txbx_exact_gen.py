# -*- coding: utf-8 -*-
"""Baseline of an EXACT-spaced line inside a text box.

blind-G ja/correspondence__14012ac2: a wps box (double 2.25pt outline) holding
ＭＳ 明朝 12pt paragraphs at line=340 exact.  With the box pinned to the page,
Word's first baseline is 13.65 below the content top (inset 3.6 + outline
1.125) = 0.8 x 17pt, Oxi's 14.32 (+0.67); every line of the box is low by the
same amount.  Body exact lines follow 0.8 L (S1261 / S1356).

Arms <face>_<size>_<line>[_nogrid]: face = msmin (ＭＳ 明朝) / cal (Calibri) /
meu (Meiryo UI); size 10.5 / 12 / 16; line twips 240 / 340 / 480 exact; default
docGrid linesAndChars 398 (the document's), nogrid = no docGrid.  Box: page
100pt, 400 x 150pt, default insets, no outline, anchor t, two paragraphs.  Read:
both baselines minus the box top (Word PDF; Oxi glyphs).

    python tools/metrics/_pb_txbx_exact_gen.py gen
    python tools/metrics/_pb_txbx_exact_gen.py word
    python tools/metrics/_pb_txbx_exact_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_txbx_exact"
EMU = 12700
TOP = 100
FACES = {"msmin": "ＭＳ 明朝", "cal": "Calibri", "meu": "Meiryo UI"}
ARMS = {}
for f in FACES:
    for s in (10.5, 12, 16):
        for ln in (240, 340, 480):
            ARMS[f"{f}_{s}_{ln}".replace(".", "p")] = (f, s, ln, True)
ARMS["msmin_12_340_nogrid"] = ("msmin", 12, 340, False)
ARMS["cal_12_340_nogrid"] = ("cal", 12, 340, False)


def box(face, size, line):
    cx, cy = 400 * EMU, 150 * EMU
    hp = int(round(size * 2))
    rpr = f'<w:rPr><w:rFonts w:ascii="{face}" w:eastAsia="{face}" w:hAnsi="{face}" w:hint="eastAsia"/><w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/></w:rPr>'
    paras = "".join(f'<w:p><w:pPr><w:spacing w:line="{line}" w:lineRule="exact"/>{rpr}</w:pPr><w:r>{rpr}<w:t>{t}</w:t></w:r></w:p>'
                    for t in ("・団体競技は各競技団体からAbc", "・個人競技は準備が整ったXyz"))
    return ('<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
            '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{TOP * EMU}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="TB1"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></wp:cNvGraphicFramePr>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
            '<wps:txbx><w:txbxContent>' + paras + '</w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
            '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>')


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (f, s, ln, grid) in ARMS.items():
        body = '<w:p>' + box(FACES[f], s, ln) + '<w:r><w:t>本文の錨段落</w:t></w:r></w:p>'
        body += "".join(f'<w:p><w:r><w:t>本文{i}行目</w:t></w:r></w:p>' for i in range(1, 4))
        g = '<w:docGrid w:type="linesAndChars" w:linePitch="398" w:charSpace="-3930"/>' if grid else ''
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>'
               '<w:cols w:space="425"/>' + g + '</w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" '
                    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def _pick(rows):
    a = [y for y, t in rows if "団体" in t]
    b = [y for y, t in rows if "個人" in t]
    return {"a": round(a[0] - TOP, 2) if a else None, "b": round(b[0] - TOP, 2) if b else None}


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
        pg = fitz.open(str(OUT / f"{arm}.pdf"))[0]; rows = []
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                rows.append((l["spans"][0]["origin"][1], "".join(s["text"] for s in l["spans"])))
        res[arm] = _pick(rows)
        print(f"{arm:20s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]; rows = {}
        for x in g:
            rows.setdefault(round(x["baseline"], 2), []).append(x["char"])
        o = _pick([(y, "".join(c)) for y, c in sorted(rows.items())])
        flag = "" if o["a"] is not None and wd[arm]["a"] is not None and abs(o["a"] - wd[arm]["a"]) < 0.3 else "  <-- differs"
        print(f"{arm:20s} oxi {o}  word {wd[arm]}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
