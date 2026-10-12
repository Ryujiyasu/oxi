# -*- coding: utf-8 -*-
"""Where does Word put the first baseline of a multiple-spaced paragraph that
follows a TABLE, for an East Asian face without a docGrid?

blind-G ja/legal__0e9cf147 (compat 15, no docGrid): the last row's «廃止年月日»
(ＭＳ ゴシック 11) is followed by «注１ 業務区域の欄には…» (ＭＳ ゴシック 9, line
276 auto). Word: row baseline 371.21 → note baseline 383.93 (12.72); Oxi
372.07 → 386.10 (14.03) — +1.3 on the one line.

Arms: <prev>_<line>_<font>: prev = tbl (1-row bordered table) / para (a plain
paragraph); line = 240 / 276 / 360; font = gothic9 (ＭＳ ゴシック 9) / mincho10
(ＭＳ 明朝 10.5).  Read: baseline(後) − baseline(前).

    python tools/metrics/_pb_aftertbl_mult_gen.py gen
    python tools/metrics/_pb_aftertbl_mult_gen.py word
    python tools/metrics/_pb_aftertbl_mult_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_aftertbl_mult"
FONTS = {"gothic9": ("ＭＳ ゴシック", 18), "mincho10": ("ＭＳ 明朝", 21)}
LINES = ("240", "276", "360")
PREVS = ("tbl", "para")
ARMS = [f"{p}_{l}_{f}" for p in PREVS for l in LINES for f in FONTS]
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl w:val="0"/></w:pPr></w:pPrDefault>'
          '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>')


def rpr(font, sz):
    return f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:hint="eastAsia"/><w:sz w:val="{sz}"/></w:rPr>'


def prev_block(prev):
    p = f'<w:p><w:pPr>{rpr("ＭＳ ゴシック", 22)}</w:pPr><w:r>{rpr("ＭＳ ゴシック", 22)}<w:t>前の行</w:t></w:r></w:p>'
    if prev == "para":
        return p
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="306" w:type="dxa"/><w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4867"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4867" w:type="dxa"/></w:tcPr>'
            + p + '</w:tc></w:tr></w:tbl>')


def doc_xml(prev, line, font):
    fam, sz = FONTS[font]
    body = (prev_block(prev) +
            f'<w:p><w:pPr><w:spacing w:line="{line}" w:lineRule="auto"/>{rpr(fam, sz)}</w:pPr><w:r>{rpr(fam, sz)}<w:t>後の行 業務区域の欄には</w:t></w:r></w:p>'
            f'<w:p><w:pPr><w:spacing w:line="{line}" w:lineRule="auto"/>{rpr(fam, sz)}</w:pPr><w:r>{rpr(fam, sz)}<w:t>次の行 用紙の大きさは</w:t></w:r></w:p>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134"'
            ' w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    for arm in ARMS:
        prev, line, font = arm.split("_")
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc_xml(prev, line, font))
    print("ok", len(ARMS))


def _read(lines):
    g = lambda k: next((l["y"] for l in lines if l["t"].startswith(k)), None)
    a, b, c = g("前の行"), g("後の行"), g("次の行")
    return {"prev_to_first": round(b - a, 2) if a and b else None, "pitch": round(c - b, 2) if b and c else None}


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
        print(f"{arm:18s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        wd = word.get(arm) or {}
        flag = "" if o["prev_to_first"] is not None and wd.get("prev_to_first") is not None and abs(o["prev_to_first"] - wd["prev_to_first"]) < 0.4 else "  <-- differs"
        print(f"{arm:18s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
