# -*- coding: utf-8 -*-
"""Where does Word split (or push) a single-column table row at the page bottom
when the cell paragraph is multiple-spaced with space-after and the table style
adds cell margins?  2026-10-13: probeqtblstyle (tblCellMar top/bottom 60 twips)
and probethdr re-measured under after=160 line=278 defaults miss by one line at
the page bottom; probecoltbl (no cell margins) now splits like Word after one line.

Arms (bottom margin swept, explicit pPrDefault after=160 line=278, ＭＳ 明朝
10.5pt, Normal widowControl=0, lines 360 grid, a 12-row single-column table of
2-line cells after nine body paragraphs (the table straddles the page bottom)):
    m0_<B>    tblCellMar top/bottom 0 (default)
    m60_<B>   tblCellMar top/bottom 60 twips (= 3pt, the probeqtblstyle style)
    m0w_<B>   m0 with widowControl ON (Normal style without widowControl 0)
Read: the number of table lines on page 1 and the last page-1 baseline
(Word PDF via fitz; Oxi via --dump-glyphs). A split row shows as an odd
line count.

    python tools/metrics/_pb_cellsplit_mult_gen.py gen
    python tools/metrics/_pb_cellsplit_mult_gen.py word
    python tools/metrics/_pb_cellsplit_mult_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402
import line_census as L  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_cellsplit_mult"
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'
CELL = "項目N：本項に定める事項については、関係法令及び本規程の趣旨に照らし、善良なる管理者の注意をもって。"
SERIES = [("m0", range(1300, 1501, 4), 0, False), ("m60", range(1300, 1501, 4), 60, False), ("m0w", range(1300, 1501, 8), 0, True)]


def styles(widow_on):
    normal = "" if widow_on else '<w:widowControl w:val="0"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault>' + RPR +
            '</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr>' + normal + '</w:pPr></w:style></w:styles>')


def para(text):
    return f'<w:p><w:pPr><w:jc w:val="both"/>{RPR}</w:pPr><w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def table(margin):
    mar = f'<w:tblCellMar><w:top w:w="{margin}" w:type="dxa"/><w:bottom w:w="{margin}" w:type="dxa"/></w:tblCellMar>' if margin else ""
    rows = "".join(f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>{para(CELL.replace("N", str(i)))}</w:tc></w:tr>' for i in range(1, 13))
    return ('<w:tbl><w:tblPr><w:tblW w:w="4500" w:type="dxa"/>' + mar +
            '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/></w:tblGrid>' + rows + '</w:tbl>')


def arms():
    for prefix, bottoms, margin, widow_on in SERIES:
        for b in bottoms:
            yield f"{prefix}_{b}", b, margin, widow_on


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    n = 0
    for arm, b, margin, widow_on in arms():
        body = "".join(para(f"第{i}条　本項に定める事項については、関係法令及び本規程の趣旨に照らし、善良なる管理者の注意をもって、誠実かつ適切に取り扱うものとする。") for i in range(1, 10)) + table(margin) + para("末尾")
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="{b}"'
               ' w:left="1418" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')
        # compatibilityMode 15 as in the golden probes (a package without it runs in legacy mode)
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:characterSpacingControl w:val="compressPunctuation"/>'
                    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles(widow_on))
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
        n += 1
    print("ok", n)


def _tbl_lines(lines):
    # cell text starts inside the cell border (x0 71.4), body text at the margin (70.9)
    t = [l for l in lines if l["x0"] > 71.2]
    return {"tbl_lines": len(t), "last_y": round(t[-1]["y"], 2) if t else None, "last_t": t[-1]["t"][:6] if t else None}


def word_p1(pdf):
    import fitz
    page = fitz.open(pdf)[0]; items = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"].strip():
                        items.append((c["origin"][1], c["origin"][0], c["bbox"][2], c["c"]))
    return _tbl_lines(L.group_lines(items))


def oxi_p1(exe, docx):
    t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
    d = json.load(open(out, encoding="utf-8"))
    items = [(g.get("baseline") or g["top"] + L.win_ascent(g.get("font_family", "")) * g["font_size"], g["x"], g["x"] + g["font_size"], g["char"]) for g in d["pages"][0]["glyphs"]]
    return _tbl_lines(L.group_lines(items))


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm, *_ in arms():
            if (OUT / f"{arm}.pdf").exists():
                continue
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm, b, *_ in arms():
        res[arm] = word_p1(OUT / f"{arm}.pdf")
        print(f"{arm:10s} bottom {b} body {round((16838 - 1418 - b) / 20, 2):6.2f}pt  word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm, b, *_ in arms():
        o = oxi_p1(os.path.abspath(exe), OUT / f"{arm}.docx")
        flag = "" if o["tbl_lines"] == wd[arm]["tbl_lines"] else "  <-- differs"
        print(f"{arm:10s} body {round((16838 - 1418 - b) / 20, 2):6.2f}pt  word {wd[arm]['tbl_lines']} last={wd[arm]['last_y']}  oxi {o['tbl_lines']} last={o['last_y']}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
