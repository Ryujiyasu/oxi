# -*- coding: utf-8 -*-
"""Does a linesAndChars grid pitch apply to full-width PUNCTUATION and parens?

`reference__0cf9c879` (ＭＳ 明朝 11pt, charSpace -2880): Word's PDF advances
'、' '。' '（' '）' at the same 10.30 grid pitch as the kanji, Oxi keeps them at
the natural 11.0 -- one '、' per line is enough to push the 44th character to the
next line, and the document to a 2nd page. Sweep the sign of charSpace and the
character class, read Word's own PDF export and group the advances by class.

    python _pb_gridpunct_gen.py gen
    python _pb_gridpunct_gen.py pdf      # Word truth (COM -> PDF -> per-class advance)
    python _pb_gridpunct_gen.py oxi      # Oxi, same arms (--dump-layout per-char widths)
"""
import collections
import json
import os
import statistics
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_gridpunct")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

MINCHO = "ＭＳ 明朝"
# every 4th character is the probe character; the last line of each paragraph is
# NOT justified, so its advances are the natural grid pitch
TEXTS = {
    "touten": "国国国、" * 60 + "国国",
    "kuten": "国国国。" * 60 + "国国",
    "paren": "国（国）" * 60 + "国国",
    "kagi": "国「国」" * 60 + "国国",
    "nakaguro": "国国国・" * 60 + "国国",
    "plain": "国" * 242,
}
ARMS = []
for cs in (-2880, 2880, None):
    for key in TEXTS:
        ARMS.append(("cs%s_%s" % ("none" if cs is None else cs, key), 22, cs, key))


def docx(label):
    return os.path.join(OUT, "gridpunct_%s.docx" % label)


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
                '<w:characterSpacingControl w:val="doNotCompress"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
    for label, sz, cs, key in ARMS:
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  "<w:docDefaults><w:rPrDefault><w:rPr>"
                  '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/><w:sz w:val="%d"/></w:rPr></w:rPrDefault>'
                  '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
                  "</w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:jc w:val="both"/></w:pPr></w:style></w:styles>' % (MINCHO, MINCHO, MINCHO, sz))
        grid = '<w:docGrid w:type="linesAndChars" w:linePitch="324"%s/>' % ("" if cs is None else ' w:charSpace="%d"' % cs)
        body = ('<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr><w:t>%s</w:t></w:r></w:p>'
                % (sz, TEXTS[key]))
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


def cls(c):
    if c == "国":
        return "kanji"
    if c.strip():
        return "probe"
    return None


def advances_pdf(path):
    import fitz
    pg = fitz.open(path)[0]
    adv = collections.defaultdict(list)
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            chars = [c for sp in l["spans"] for c in sp["chars"] if c["c"].strip()]
            if len(chars) < 3:
                continue
            xs = [c["origin"][0] for c in chars]
            full = xs[-1] > 500.0
            for i in range(len(chars) - 1):
                k = cls(chars[i]["c"])
                if k:
                    adv[(k, "full" if full else "last")].append(round(xs[i + 1] - xs[i], 2))
    return adv


def report(adv, label):
    row = []
    for k in (("kanji", "last"), ("probe", "last"), ("kanji", "full"), ("probe", "full")):
        v = adv.get(k)
        row.append("%s %s=%s" % (k[0][:5], k[1], ("%.2f(n%d)" % (statistics.median(v), len(v))) if v else "-"))
    print("%-20s %s" % (label, "  ".join(row)))


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for label, _, _, _ in ARMS:
            out = docx(label)[:-5] + ".word.pdf"
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                d.SaveAs2(out, 17)
            finally:
                d.Close(False)
    finally:
        app.Quit()
    print("== WORD (PDF export, median advance pt; 'last' = unjustified last line) ==")
    for label, _, _, _ in ARMS:
        report(advances_pdf(docx(label)[:-5] + ".word.pdf"), label)


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    print("== OXI %s (dump widths, median pt) ==" % (envs or "(default)"))
    for label, _, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "gridpunct_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "gpu"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        adv = collections.defaultdict(list)
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            rows = collections.defaultdict(list)
            for e in pg["elements"]:
                if e["type"] == "text" and (e.get("text") or "").strip():
                    rows[round(e["y"], 1)].append(e)
            for y, es in rows.items():
                es.sort(key=lambda e: e["x"])
                full = es[-1]["x"] + es[-1]["w"] > 500.0
                for e in es:
                    for c in e["text"]:
                        k = cls(c)
                        if k:
                            adv[(k, "full" if full else "last")].append(round(e["w"], 2))
        report(adv, label)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
