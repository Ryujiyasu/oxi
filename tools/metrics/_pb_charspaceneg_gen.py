# -*- coding: utf-8 -*-
"""What character pitch does a linesAndChars grid with a NEGATIVE charSpace give?

`reference__0cf9c879` (ＭＳ 明朝 11pt, `<w:docGrid w:type="linesAndChars"
w:linePitch="324" w:charSpace="-2880"/>`) holds 44 characters per 453pt line in
Word (pitch 10.30) where Oxi holds 41 (pitch 11.0 = the font size, i.e. the
negative value ignored). Sweep charSpace and the font size and read the pitch
as line width / characters per line over a long paragraph of one CJK
character (COM `ComputeStatistics(wdStatisticLines)` gives the line count).

    python _pb_charspaceneg_gen.py gen
    python _pb_charspaceneg_gen.py pdf      # Word truth (COM)
    python _pb_charspaceneg_gen.py oxi      # Oxi, same arms (dump line count)
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_charspaceneg")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

MINCHO = "ＭＳ 明朝"
NCHARS = 600
WIDTH_PT = (11906 - 2 * 1418) / 20.0      # 453.0

# (label, sz half-points, charSpace or None, docGrid type)
ARMS = []
for sz in (22, 21, 24):
    for cs in (None, 0, -1440, -2880, -4320, 1440, 2880):
        ARMS.append(("sz%d_cs%s" % (sz, "none" if cs is None else cs), sz, cs, "linesAndChars"))
ARMS.append(("sz22_lines_cs-2880", 22, -2880, "lines"))


def docx(label):
    return os.path.join(OUT, "charspaceneg_%s.docx" % label)


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
    for label, sz, cs, gtype in ARMS:
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  "<w:docDefaults><w:rPrDefault><w:rPr>"
                  '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/><w:sz w:val="%d"/></w:rPr></w:rPrDefault>'
                  '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
                  "</w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:jc w:val="both"/></w:pPr></w:style></w:styles>' % (MINCHO, MINCHO, MINCHO, sz))
        grid = '<w:docGrid w:type="%s" w:linePitch="324"%s/>' % (gtype, "" if cs is None else ' w:charSpace="%d"' % cs)
        body = ('<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr><w:t>%s</w:t></w:r></w:p>'
                % (sz, "国" * NCHARS))
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


def report(nlines, who):
    print("== %s ==" % who)
    print("%-22s %-5s %-9s %-6s %-9s %-9s %s" % ("arm", "pt", "charSpace", "lines", "chars/ln", "pitch", "pitch - pt"))
    for label, sz, cs, gtype in ARMS:
        n = nlines.get(label)
        if not n:
            print("%-22s %-5.1f %-9s -" % (label, sz / 2.0, cs)); continue
        cpl = NCHARS / n
        # the last line is partial; the full lines hold floor(WIDTH/pitch) chars
        full = (NCHARS - (NCHARS % max(1, int(cpl)))) / (n - 1) if n > 1 else cpl
        pitch = WIDTH_PT / full
        print("%-22s %-5.1f %-9s %-6d %-9.2f %-9.3f %+.3f" % (label, sz / 2.0, cs, n, full, pitch, pitch - sz / 2.0))


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    nl = {}
    try:
        for label, _, _, _ in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                nl[label] = d.Paragraphs(1).Range.ComputeStatistics(1)
            finally:
                d.Close(False)
    finally:
        app.Quit()
    report(nl, "WORD (COM line count)")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    nl = {}
    for label, _, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "charspaceneg_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "csn"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        ys = set()
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and (e.get("text") or "").strip():
                    ys.add((pg.get("index", 0), round(e["y"], 1)))
        nl[label] = len(ys)
    report(nl, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
