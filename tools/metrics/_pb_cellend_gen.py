# -*- coding: utf-8 -*-
"""Does the EMPTY end-of-cell paragraph add a line to the row?

`correspondence__04a3e3e1`'s timetable (HG丸ｺﾞｼｯｸM-PRO, `adjustLineHeightInTable`,
`snapToGrid=0` paragraphs): every content cell ends with an empty paragraph that
names no size (10pt default). Word's PDF puts the row's bottom rule at that
paragraph's Info6 (503.23 vs 503.25) and the row is exactly the tallest cell's
TEXT height -- the cell-end mark takes ~0. Oxi counts it at its natural line
(10 x 83/64 = 12.97), which is where S1310's +13/row came from.

Arms stack [基準] [one-cell table] [次] [末尾]; the table row is read as the span
基準 -> 次 minus the same span of the `notable` control, so the row height is
what is compared (Word quantises Info6 to 0.75pt; a 13pt line is unmistakable).
Cell contents vary the number of text lines, whether an empty paragraph ends
the cell, its size (named 9pt vs unnamed 10pt) and the ALIT compat flag.

    python _pb_cellend_gen.py gen
    python _pb_cellend_gen.py pdf      # Word truth (COM Info6)
    python _pb_cellend_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cellend")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

FACE = "HG丸ｺﾞｼｯｸM-PRO"
# cell content recipes: 't' = 9pt text line, 'e9' = empty 9pt, 'e10' = empty, no size
RECIPES = {
    "t1": ["t"], "t1_e10": ["t", "e10"], "t1_e9": ["t", "e9"], "t1_e9_e10": ["t", "e9", "e10"],
    "t3": ["t", "t", "t"], "t3_e10": ["t", "t", "t", "e10"], "t3_e9_e10": ["t", "t", "t", "e9", "e10"],
    "e10": ["e10"], "e9_e10": ["e9", "e10"], "e9": ["e9"],
}
ARMS = [("notable_alit0", None, False), ("notable_alit1", None, True)]
for r in RECIPES:
    ARMS.append(("%s_alit0" % r, r, False))
    ARMS.append(("%s_alit1" % r, r, True))


def docx(label):
    return os.path.join(OUT, "cellend_%s.docx" % label)


def para(kind):
    fonts = '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:cs="ＭＳ Ｐゴシック"/>' % (FACE, FACE, FACE)
    if kind == "t":
        rpr = fonts + '<w:sz w:val="18"/><w:szCs w:val="18"/>'
        return ('<w:p><w:pPr><w:widowControl/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/>'
                '<w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:t>1:00～2:00</w:t></w:r></w:p>' % (rpr, rpr))
    rpr = fonts + ('<w:sz w:val="18"/><w:szCs w:val="18"/>' if kind == "e9" else "")
    return ('<w:p><w:pPr><w:widowControl/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/>'
            '<w:rPr>%s</w:rPr></w:pPr></w:p>' % rpr)


def table(recipe):
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
            '<w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>'
            '<w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders>'
            '<w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar>'
            '</w:tblPr><w:tblGrid><w:gridCol w:w="1694"/></w:tblGrid>'
            '<w:tr><w:trPr><w:trHeight w:val="63"/></w:trPr><w:tc><w:tcPr>'
            '<w:tcW w:w="1694" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>'
            % "".join(para(k) for k in RECIPES[recipe]))


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
    # the witness's compat block, minus/plus adjustLineHeightInTable
    compat = ('<w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/>'
              '<w:ulTrailSpace/><w:doNotExpandShiftReturn/>%s<w:useFELayout/>'
              '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
              '<w:jc w:val="both"/></w:pPr></w:style></w:styles>')
    for label, recipe, alit in ARMS:
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                    "<w:compat>" + compat % ("<w:adjustLineHeightInTable/>" if alit else "") + "</w:compat>"
                    '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
        body = marker("基準") + (table(recipe) if recipe else "") + marker("次") + marker("末尾")
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
    print("%-16s %-6s %-9s %-9s %s" % ("recipe", "alit", "span", "row", "row - text-only row"))
    for label, recipe, alit in ARMS:
        if recipe is None:
            continue
        sp = spans.get(label)
        ctl = spans.get("notable_alit%d" % alit)
        row = None if (sp is None or ctl is None) else sp - ctl
        base_label = recipe.split("_")[0] + "_alit%d" % alit          # t1 / t3 / e10 / e9 ...
        base = spans.get(base_label)
        d = None if (row is None or base is None or ctl is None) else row - (base - ctl)
        print("%-16s %-6s %-9s %-9s %s" % (recipe, "on" if alit else "-",
                                         "-" if sp is None else "%.2f" % sp,
                                         "-" if row is None else "%.2f" % row,
                                         "" if d is None or base_label == label else "%+.2f" % d))


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
    for label, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "cellend_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "ce"),
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
