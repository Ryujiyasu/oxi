# -*- coding: utf-8 -*-
"""A run of line-start-prohibited characters (katakana middle dots) overflows
the line by k cells: does Word break before the run (pull-back), break inside
it, or let the line overhang?

blind-G ja/reports__58dca3b2 (linesAndChars charSpace=-2271, 2 columns, 5pt
HG丸ｺﾞｼｯｸM-PRO): «・・・き り と り» + 56 «・» = 65 cells of 4.44 = 288.6pt in a
283.5pt column; Word keeps ONE line ending at 288.6 (overhang 1.15 cells);
Oxi breaks before the «り» that precedes the dots and carries 57 characters
to a second line, which then overflows into the next column.

One doc per total-cell count N in {56, 57, 58, 59, 60, 62, 70} (one column of
566.9pt, 10.5pt cells of 9.95 → 56.97 cells fit): «あいう» + (N − 3) «・», with
a sibling family where the dots are «。» (a hanging mark) for contrast.
Read: number of lines and line-1 x1.

    python tools/metrics/_pb_dotrun_overflow_gen.py gen
    python tools/metrics/_pb_dotrun_overflow_gen.py word
    python tools/metrics/_pb_dotrun_overflow_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_dotrun_overflow"
NS_ = (56, 57, 58, 59, 60, 62, 70, 90, 115, 120, 150)
MARKS = {"dot": "・", "maru": "。"}
ARMS = [f"{m}_{n}" for m in MARKS for n in NS_]
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="HG丸ｺﾞｼｯｸM-PRO" w:eastAsia="HG丸ｺﾞｼｯｸM-PRO" w:hAnsi="HG丸ｺﾞｼｯｸM-PRO" w:cs="Times New Roman"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="840"/>'
                '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    for arm in ARMS:
        m, n = arm.split("_"); n = int(n)
        text = "あいう" + MARKS[m] * (n - 3)
        body = (f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'
                f'<w:p><w:r><w:t>次の段落</w:t></w:r></w:p>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="454" w:right="284" w:bottom="284" w:left="284"'
               ' w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="linesAndChars" w:linePitch="290" w:charSpace="-2271"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def _read(lines):
    ls = [l for l in lines if not l["t"].startswith("次の段落")]
    ls.sort(key=lambda l: l["y"])
    return {"lines": len(ls), "x1": [round(l["x1"], 1) for l in ls], "n1": len(ls[0]["t"]) if ls else None}


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
        print(f"{arm:9s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        wd = word.get(arm) or {}
        flag = "" if o["lines"] == wd.get("lines") and o["n1"] == wd.get("n1") else "  <-- differs"
        print(f"{arm:9s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
