# -*- coding: utf-8 -*-
"""Does a line-spacing MULTIPLIER compete with the grid the same way whatever
its ORIGIN -- direct paragraph property, paragraph style, or docDefaults?

`reference__0cf9c879` (ＭＳ 明朝 11pt, linesAndChars linePitch 324 = 16.2pt,
`pPrDefault` `line=276 auto`) advances every body line exactly 16.275 in Word:
one cell, no x1.15. The S1306 law (`pitch x max(cells, mult)`) gives 18.63 and
that is what Oxi does. S1306's probe put the multiplier on the paragraph
itself; this sweeps where it comes from.

Each arm stacks three one-line paragraphs between 基準 and 次; the per-line
pitch is (span - control span) / 3 + control's own line... simpler: report
the span itself and compare arms (the control has x1.0 everywhere).

    python _pb_gridmultsrc_gen.py gen
    python _pb_gridmultsrc_gen.py pdf
    python _pb_gridmultsrc_gen.py oxi
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_gridmultsrc")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

MINCHO = "ＭＳ 明朝"
N = 3
# (label, multiplier source: none / direct / style / docdefaults, w:line, sz half-points, grid)
LC = '<w:docGrid w:type="linesAndChars" w:linePitch="324" w:charSpace="-2880"/>'
LC0 = '<w:docGrid w:type="linesAndChars" w:linePitch="324"/>'
LN = '<w:docGrid w:type="lines" w:linePitch="324"/>'
ARMS = [
    ("ctl_x100", "none", 240, 22, LC),
    ("direct_x115", "direct", 276, 22, LC),
    ("style_x115", "style", 276, 22, LC),
    ("docdef_x115", "docdefaults", 276, 22, LC),
    ("docdef_x115_charspace0", "docdefaults", 276, 22, LC0),
    ("docdef_x115_lines", "docdefaults", 276, 22, LN),
    ("direct_x150", "direct", 360, 22, LC),
    ("docdef_x150", "docdefaults", 360, 22, LC),
    # 10.5pt: natural 13.6 < 16.2 as well
    ("ctl_x100_21", "none", 240, 21, LC),
    ("direct_x115_21", "direct", 276, 21, LC),
    ("docdef_x115_21", "docdefaults", 276, 21, LC),
    # 14pt: natural 18.2 > 16.2 -> 2 cells; does the origin matter there?
    ("ctl_x100_28", "none", 240, 28, LC),
    ("direct_x115_28", "direct", 276, 28, LC),
    ("docdef_x115_28", "docdefaults", 276, 28, LC),
]


def docx(label):
    return os.path.join(OUT, "gridmultsrc_%s.docx" % label)


def para(sz, direct_line, style_id):
    ppr = ""
    if direct_line is not None:
        ppr += '<w:spacing w:line="%d" w:lineRule="auto"/>' % direct_line
    if style_id:
        ppr = '<w:pStyle w:val="%s"/>' % style_id + ppr
    return ('<w:p><w:pPr>%s<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
            '<w:sz w:val="%d"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
            '<w:sz w:val="%d"/></w:rPr><w:t>本文の行送りを測る</w:t></w:r></w:p>'
            % (ppr, MINCHO, MINCHO, MINCHO, sz, MINCHO, MINCHO, MINCHO, sz))


def marker(text):
    return ('<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/><w:sz w:val="22"/></w:rPr>'
            '<w:t>%s</w:t></w:r></w:p>' % (MINCHO, MINCHO, MINCHO, text))


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
                '<w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
    for label, src, line, sz, grid in ARMS:
        ppr_default = ('<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="%d" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
                       % (line if src == "docdefaults" else 240))
        style_extra = ""
        if src == "style":
            style_extra = ('<w:style w:type="paragraph" w:styleId="b"><w:name w:val="mult"/><w:basedOn w:val="a"/>'
                           '<w:pPr><w:spacing w:line="%d" w:lineRule="auto"/></w:pPr></w:style>' % line)
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  "<w:docDefaults><w:rPrDefault><w:rPr>"
                  '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
                  "%s</w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>%s</w:styles>'
                  % (MINCHO, MINCHO, MINCHO, ppr_default, style_extra))
        p = para(sz, line if src == "direct" else None, "b" if src == "style" else None)
        body = marker("基準") + p * N + marker("次") + marker("末尾")
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>'
               + grid + "</w:sectPr></w:body></w:document>")
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
    print("%-26s %-12s %-5s %-5s %-9s %s" % ("arm", "mult source", "pt", "mult", "span", "per test line (span-ctl)/3 + ctl line"))
    for label, src, line, sz, grid in ARMS:
        sp = spans.get(label)
        ctl = spans.get("ctl_x100" if sz == 22 else "ctl_x100_%d" % sz)
        per = "" if (sp is None or ctl is None) else "%+.2f vs control" % ((sp - ctl) / N)
        print("%-26s %-12s %-5.1f %-5.2f %-9s %s" % (label, src, sz / 2.0, line / 240.0, "-" if sp is None else "%.2f" % sp, per))


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    spans = {}
    try:
        for label, _, _, _, _ in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                ys = {}
                for i in range(1, d.Paragraphs.Count + 1):
                    p = d.Paragraphs(i)
                    st = d.Range(p.Range.Start, p.Range.Start)
                    ys.setdefault((p.Range.Text or "").rstrip("\r\x07"), float(st.Information(6)))
                spans[label] = ys["次"] - ys["基準"]
            finally:
                d.Close(False)
    finally:
        app.Quit()
    report(spans, "WORD (Info6, collapsed starts)")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    spans = {}
    for label, _, _, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "gridmultsrc_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "gm"),
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
