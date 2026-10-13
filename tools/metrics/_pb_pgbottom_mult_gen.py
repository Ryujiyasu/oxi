# -*- coding: utf-8 -*-
"""Where exactly does Word stop filling a page when the document defaults are
after=160 line=278 (auto) on a lines grid?  2026-10-13: the pPrDefault-less
golden probes re-measured under this machine's Normal.dotm (after=160 line=278)
differ from Oxi by ONE line at the page bottom in both directions
(probelac / probefn / probecoltbl / probeqtblstyle ...), with every in-page
pitch matching (pitch = grid pitch x 278/240).

Arms: the bottom margin is swept (see SERIES): single_<B> one long paragraph
on a lines-360 grid in 1-twip steps (the exact allowance), para_<B> 14
three-line paragraphs (a paragraph START after an 8pt after), sz24_<B> 12pt
text, nogrid_<B> no docGrid, p400_<B> a 20pt grid pitch. Explicit pPrDefault
(after=160 line=278), ＭＳ 明朝, Normal widowControl=0, jc both (as the golden
probes), top margin 1418.
Read: number of lines on page 1 and the last page-1 baseline (Word PDF via
fitz; Oxi via --dump-glyphs).

    python tools/metrics/_pb_pgbottom_mult_gen.py gen
    python tools/metrics/_pb_pgbottom_mult_gen.py word
    python tools/metrics/_pb_pgbottom_mult_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402
import line_census as L  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_pgbottom_mult"
# series: (prefix, bottoms, grid xml, sz half-points, paragraphs)
SERIES = [
    ("single", range(1300, 1321, 1), '<w:docGrid w:type="lines" w:linePitch="360"/>', 21, False),   # 1-twip steps: the exact page-bottom allowance
    ("para", range(800, 1301, 10), '<w:docGrid w:type="lines" w:linePitch="360"/>', 21, True),      # a paragraph START after an 8pt after
    ("sz24", range(1200, 1401, 4), '<w:docGrid w:type="lines" w:linePitch="360"/>', 24, False),     # 12pt text, same grid
    ("nogrid", range(1300, 1501, 4), '', 21, False),                                                 # no docGrid: natural x 278/240
    ("p400", range(1300, 1501, 4), '<w:docGrid w:type="lines" w:linePitch="400"/>', 21, False),     # 20pt grid pitch
]
TEXT = ("本項に定める事項については、関係法令及び本規程の趣旨に照らし、善良なる管理者の注意をもって、誠実かつ適切に取り扱うものとし、"
        "これに関連して生じた一切の事項は、別に定めるところにより、遅滞なく、かつ、公正に処理しなければならない。")
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault>' + RPR +
          '</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>'
          '</w:styles>')


def para(text, rpr=RPR):
    return f'<w:p><w:pPr><w:jc w:val="both"/>{rpr}</w:pPr><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def arms():
    for prefix, bottoms, grid, sz, paras in SERIES:
        rpr = RPR.replace('w:val="21"', f'w:val="{sz}"')
        for b in bottoms:
            body = "".join(para(f"第{i}条　" + TEXT, rpr) for i in range(1, 15)) if paras else para("　".join(TEXT for _ in range(14)), rpr)
            yield f"{prefix}_{b}", b, body, grid, rpr


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    n = 0
    for arm, b, body, grid, rpr in arms():
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="{b}"'
               ' w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>' + grid + '</w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", T.CT); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", T.DRELS); z.writestr("word/styles.xml", STYLES.replace(RPR, rpr))
            z.writestr("word/document.xml", doc)
        n += 1
    print("ok", n)


def word_p1(pdf):
    import fitz
    page = fitz.open(pdf)[0]
    items = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"].strip():
                        items.append((c["origin"][1], c["origin"][0], c["bbox"][2], c["c"]))
    lines = L.group_lines(items)
    return {"n": len(lines), "last_y": round(lines[-1]["y"], 2) if lines else None, "first_y": round(lines[0]["y"], 2) if lines else None}


def oxi_p1(exe, docx):
    t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
    d = json.load(open(out, encoding="utf-8"))
    items = []
    for g in d["pages"][0]["glyphs"]:
        fs = g.get("font_size", 0) or 0
        base = g.get("baseline")
        if base is None:
            base = g["top"] + L.win_ascent(g.get("font_family", "")) * fs
        items.append((base, g["x"], g["x"] + fs, g["char"]))
    lines = L.group_lines(items)
    return {"n": len(lines), "last_y": round(lines[-1]["y"], 2) if lines else None, "first_y": round(lines[0]["y"], 2) if lines else None}


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
        print(f"{arm:12s} bottom {b} body {round((16838 - 1418 - b) / 20, 2):6.2f}pt  word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm, b, *_ in arms():
        o = oxi_p1(os.path.abspath(exe), OUT / f"{arm}.docx")
        flag = "" if o["n"] == wd[arm]["n"] else "  <-- differs"
        print(f"{arm:12s} body {round((16838 - 1418 - b) / 20, 2):6.2f}pt  word n={wd[arm]['n']} last={wd[arm]['last_y']}  oxi n={o['n']} last={o['last_y']}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
