# -*- coding: utf-8 -*-
"""How much does a RUBY run raise a line -- per paragraph, or per LINE that
carries one -- and does the rule hold inside a table cell?

`correspondence__04a3e3e1`'s timetable cells (HG丸ｺﾞｼｯｸM-PRO 10pt, ruby hps 10 /
hpsRaise 18 / hpsBaseText 20): Word's lines without ruby pitch 12.75-13.5
(= 12.97 natural), lines WITH ruby 16.5-17.25, and a ruby paragraph wrapped to
three lines is 50.25 = 3 x 16.75 -- every wrapped line grows, not just one.
Oxi gives every cell line 12.97: the body path's `paragraph_ruby_expansion_pt`
(added once, at the last-line advance) never reaches the cell path, and even in
the body it is once per paragraph.

Arms: body and cell x {no ruby; one ruby line; a paragraph wrapped to 3 lines
with ruby on EVERY line; the same with ruby only on the first line}. Each test
paragraph sits between 基準 and 次; the value read is the span minus the arm's
no-ruby control, so what is compared is the ruby's ADDED height.

    python _pb_cellruby_gen.py gen
    python _pb_cellruby_gen.py pdf      # Word truth (COM Info6)
    python _pb_cellruby_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cellruby")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

FACE = "HG丸ｺﾞｼｯｸM-PRO"
FONTS = '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:cs="ＭＳ Ｐゴシック"/>' % (FACE, FACE, FACE)
RPR = FONTS + '<w:sz w:val="20"/><w:szCs w:val="20"/>'


def run(text):
    return '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % (RPR, text)


def ruby(base, rt):
    return ('<w:r><w:rPr>%s</w:rPr><w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/>'
            '<w:hps w:val="10"/><w:hpsRaise w:val="18"/><w:hpsBaseText w:val="20"/><w:lid w:val="ja-JP"/>'
            '</w:rubyPr><w:rt><w:r><w:rPr>%s<w:sz w:val="10"/></w:rPr><w:t>%s</w:t></w:r></w:rt>'
            '<w:rubyBase><w:r><w:rPr>%s</w:rPr><w:t>%s</w:t></w:r></w:rubyBase></w:ruby></w:r>'
            % (RPR, FONTS, rt, RPR, base))


# paragraph bodies (cell width 84.7pt holds ~8 chars of 10pt)
PLAIN1 = run("1:00～2:00")
RUBY1 = ruby("正午", "しょうご") + run("～7:00")
# three lines, ruby on every line: 年金/相談会 ... 年金 ... 主催
RUBY3 = ruby("年金", "ねんきん") + ruby("相談会", "そうだんかい") + run("（") + ruby("年金", "ねんきん") + run("ﾄｰﾀﾙｻﾎﾟｰﾄ･ｺｽﾓ") + ruby("主催", "しゅさい") + run("）")
# three lines, ruby only on the first
RUBY3_FIRST = ruby("年金", "ねんきん") + run("相談会（年金ﾄｰﾀﾙｻﾎﾟｰﾄ･ｺｽﾓ主催）")
PLAIN3 = run("年金相談会（年金ﾄｰﾀﾙｻﾎﾟｰﾄ･ｺｽﾓ主催）")

# (label, in cell?, paragraph content)
ARMS = [
    ("body_plain1", False, PLAIN1), ("body_ruby1", False, RUBY1),
    ("body_plain3", False, PLAIN3), ("body_ruby3", False, RUBY3), ("body_ruby3first", False, RUBY3_FIRST),
    ("cell_plain1", True, PLAIN1), ("cell_ruby1", True, RUBY1),
    ("cell_plain3", True, PLAIN3), ("cell_ruby3", True, RUBY3), ("cell_ruby3first", True, RUBY3_FIRST),
]
CONTROL = {"ruby1": "plain1", "ruby3": "plain3", "ruby3first": "plain3"}


def docx(label):
    return os.path.join(OUT, "cellruby_%s.docx" % label)


def para(inner):
    return ('<w:p><w:pPr><w:widowControl/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/>'
            '<w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (RPR, inner))


def table(inner):
    return ('<w:tbl><w:tblPr><w:tblW w:w="1694" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>'
            '<w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>'
            '<w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders>'
            '<w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar>'
            '</w:tblPr><w:tblGrid><w:gridCol w:w="1694"/></w:tblGrid>'
            '<w:tr><w:trPr><w:trHeight w:val="63"/></w:trPr><w:tc><w:tcPr>'
            '<w:tcW w:w="1694" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>' % inner)


def marker(text):
    return ('<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:hint="eastAsia"/>'
            '<w:sz w:val="24"/></w:rPr><w:t>%s</w:t></w:r></w:p>' % text)


def gen():
    os.makedirs(OUT, exist_ok=True)
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                    "</Types>")
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
             'relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
             'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
             'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             "</Relationships>")
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                '<w:compat><w:adjustLineHeightInTable/><w:useFELayout/>'
                '<w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
              '<w:jc w:val="both"/></w:pPr></w:style></w:styles>')
    for label, in_cell, inner in ARMS:
        p = para(inner)
        # the body arms get the cell's 84.7pt line width through indents so the
        # three-line paragraphs wrap the same way
        if not in_cell:
            p = p.replace('<w:snapToGrid w:val="0"/>', '<w:snapToGrid w:val="0"/><w:ind w:right="7838"/>')
        body = marker("基準") + (table(p) if in_cell else p) + marker("次") + marker("末尾")
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
                 '<w:docGrid w:type="lines" w:linePitch="344"/>'
                 "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def report(spans, who):
    print("== %s ==" % who)
    print("%-18s %-9s %-9s %s" % ("arm", "span", "control", "ruby added"))
    for label, _, _ in ARMS:
        sp = spans.get(label)
        kind = label.split("_", 1)[1]
        ctl_label = label.split("_", 1)[0] + "_" + CONTROL.get(kind, kind)
        ctl = spans.get(ctl_label)
        d = None if (sp is None or ctl is None or kind not in CONTROL) else sp - ctl
        print("%-18s %-9s %-9s %s" % (label, "-" if sp is None else "%.2f" % sp,
                                     "-" if ctl is None else "%.2f" % ctl, "" if d is None else "%+.2f" % d))


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    spans = {}
    try:
        for label, _, _ in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                ys = {}
                for i in range(1, d.Paragraphs.Count + 1):
                    p = d.Paragraphs(i)
                    st = d.Range(p.Range.Start, p.Range.Start)
                    ys.setdefault((p.Range.Text or "").rstrip("\r\x07"), float(st.Information(6)))
                spans[label] = ys["次"] - ys["基準"]
                n_lines = d.Paragraphs(2).Range.ComputeStatistics(1)  # wdStatisticLines
                spans[label + "__lines"] = n_lines
            finally:
                d.Close(False)
    finally:
        app.Quit()
    report(spans, "WORD (Info6, collapsed starts)")
    print("   Word line counts:", {k[:-7]: v for k, v in spans.items() if k.endswith("__lines")})


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    spans = {}
    for label, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "cellruby_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "cr"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        by_y = {}
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and (e.get("text") or "").strip():
                    by_y.setdefault(round(e["y"], 2), []).append((e["x"], e["text"]))
        y = {}
        for yy, frags in sorted(by_y.items()):
            t = "".join(t for _, t in sorted(frags)).strip()
            for key in ("基準", "次"):
                if t.startswith(key):
                    y.setdefault(key, yy)
        if "基準" in y and "次" in y:
            spans[label] = y["次"] - y["基準"]
    report(spans, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
