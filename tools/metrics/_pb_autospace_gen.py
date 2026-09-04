# -*- coding: utf-8 -*-
"""Auto-space between digits/Latin and CJK (autoSpaceDN / autoSpaceDE): when does
Word insert it, and how much, in the body and in a table cell, with and without
`balanceSingleByteDoubleByteWidth`?

`correspondence__04a3e3e1`'s cell line '(資料代50円)' shows NO gap in Word's PDF on
either side of '50' (代->5 = 10.2, 0->円 = 7.3, i.e. the bare advances) while Oxi
adds 2.5pt after the digits. That 2.5 is what tips the line once the ruby
spread (S1314) is right. Each arm is one paragraph; the value read is the gap
between the last CJK char before the digits and the first digit, and between
the last digit and the next CJK char, from the PDF glyph origins (Word) and
the dump (Oxi).

    python _pb_autospace_gen.py gen
    python _pb_autospace_gen.py pdf
    python _pb_autospace_gen.py oxi
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_autospace")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

FACE = "HG丸ｺﾞｼｯｸM-PRO"
TEXTS = {"paren": ("(資料代", "50", "円)"), "plain": ("資料代", "50", "円です"), "latin": ("資料", "ABC", "円です")}
# (label, in cell?, text key, balance compat?, autoSpaceDN override (None / "0"))
ARMS = []
for cell in (False, True):
    for tk in ("paren", "plain", "latin"):
        for bal in (False, True):
            for asdn in (None, "0"):
                ARMS.append(("%s_%s_bal%d_as%s" % ("cell" if cell else "body", tk, int(bal), "def" if asdn is None else asdn), cell, tk, bal, asdn))
# the witness's digit runs carry w:hint="eastAsia": are hinted digits East Asian for autospace?
RUBY_ARMS = [("body_rubyboth", False, "paren", True, None, "both"), ("cell_rubyboth", True, "paren", True, None, "both"),
             ("body_rubyleft", False, "paren", True, None, "left"), ("body_rubyright", False, "paren", True, None, "right"),
             ("cell_rubyleft", True, "paren", True, None, "left"), ("cell_rubyright", True, "paren", True, None, "right")]
HINT_ARMS = [("body_paren_hint", False, "paren", True, None, True), ("cell_paren_hint", True, "paren", True, None, True),
             ("body_plain_hint", False, "plain", False, None, True), ("cell_latin_hint", True, "latin", True, None, True)]
ARMS = [a + (False,) for a in ARMS] + HINT_ARMS + RUBY_ARMS


def docx(label):
    return os.path.join(OUT, "autospace_%s.docx" % label)


def run(text, hint=False):
    return ('<w:r><w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:cs="ＭＳ Ｐゴシック"%s/>'
            '<w:sz w:val="20"/></w:rPr><w:t>%s</w:t></w:r>' % (FACE, FACE, FACE, ' w:hint="eastAsia"' if hint else "", text))


def ruby_run(base, rt):
    f = '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:cs="ＭＳ Ｐゴシック"/>' % (FACE, FACE, FACE)
    return ('<w:r><w:rPr>%s<w:sz w:val="20"/></w:rPr><w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="10"/>'
            '<w:hpsRaise w:val="18"/><w:hpsBaseText w:val="20"/><w:lid w:val="ja-JP"/></w:rubyPr><w:rt><w:r><w:rPr>%s<w:sz w:val="10"/></w:rPr>'
            '<w:t>%s</w:t></w:r></w:rt><w:rubyBase><w:r><w:rPr>%s<w:sz w:val="20"/></w:rPr><w:t>%s</w:t></w:r></w:rubyBase></w:ruby></w:r>' % (f, f, rt, f, base))


def para(tk, asdn, hint=False):
    a, b, c = TEXTS[tk]
    ppr = '<w:widowControl/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/>'
    if asdn is not None:
        ppr += '<w:autoSpaceDE w:val="%s"/><w:autoSpaceDN w:val="%s"/>' % (asdn, asdn)
    if hint in ("both", "left", "right"):
        left = run("(") + (ruby_run("資料代", "しりょうだい") if hint in ("both", "left") else run("資料代"))
        right = (ruby_run("円", "えん") if hint in ("both", "right") else run("円")) + run(")")
        return "<w:p><w:pPr>%s</w:pPr>%s%s%s</w:p>" % (ppr, left, run(b, True), right)
    return "<w:p><w:pPr>%s</w:pPr>%s%s%s</w:p>" % (ppr, run(a, hint), run(b, hint), run(c, hint))


def table(inner):
    return ('<w:tbl><w:tblPr><w:tblW w:w="2932" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2932"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2932" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>' % inner)


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
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
    for label, cell, tk, bal, asdn, hint in ARMS:
        compat = ('<w:spaceForUL/>%s<w:doNotLeaveBackslashAlone/><w:ulTrailSpace/><w:doNotExpandShiftReturn/>'
                  '<w:adjustLineHeightInTable/><w:useFELayout/>' % ("<w:balanceSingleByteDoubleByteWidth/>" if bal else ""))
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                    "<w:compat>" + compat + '<w:compatSetting w:name="compatibilityMode"'
                    ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                    '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
        p = para(tk, asdn, hint)
        body = (table(p) if cell else p)
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


def gaps_from_positions(chars, tk):
    """chars: list of (x, w, ch) sorted by x. Return (gap before digits, gap after digits)."""
    a, b, c = TEXTS[tk]
    txt = "".join(ch for _, _, ch in chars)
    i = txt.find(b)
    if i < 1 or i + len(b) >= len(chars):
        return None
    before = chars[i][0] - (chars[i - 1][0] + chars[i - 1][1])
    j = i + len(b)
    after = chars[j][0] - (chars[j - 1][0] + chars[j - 1][1])
    return (round(before, 2), round(after, 2))


def report(res, who):
    print("== %s ==" % who)
    print("%-26s %-5s %-6s %-4s %-6s %-5s %s" % ("arm", "cell", "text", "bal", "asDN", "hint", "gap before / after digits"))
    for label, cell, tk, bal, asdn, hint in ARMS:
        g = res.get(label)
        print("%-26s %-5s %-6s %-4d %-6s %-5s %s" % (label, "cell" if cell else "body", tk, int(bal), "def" if asdn is None else asdn,
                                                    (hint if isinstance(hint, str) else ("ea" if hint else "-")), "-" if g is None else "%+.2f / %+.2f" % g))


def pdf():
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    res = {}
    try:
        for label, cell, tk, bal, asdn, hint in ARMS:
            src, dst = docx(label), docx(label).replace(".docx", ".pdf")
            d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            try:
                d.ExportAsFixedFormat(dst, 17)
            finally:
                d.Close(False)
            pg = fitz.open(dst)[0]
            chars = []
            for b in pg.get_text("rawdict")["blocks"]:
                for ln in b.get("lines", []):
                    for s in ln["spans"]:
                        for ch in s["chars"]:
                            if ch["c"].strip():
                                chars.append((ch["bbox"][0], ch["bbox"][2] - ch["bbox"][0], ch["c"]))
            chars.sort()
            res[label] = gaps_from_positions(chars, tk)
    finally:
        app.Quit()
    report(res, "WORD (PDF glyph origins)")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    res = {}
    for label, cell, tk, bal, asdn, hint in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "autospace_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "as"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        chars = []
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and (e.get("text") or "").strip():
                    # a fragment may hold several chars; split its width evenly per char
                    t = e["text"]; n = len(t)
                    for k, ch in enumerate(t):
                        chars.append((e["x"] + e.get("w", 0) * k / n, e.get("w", 0) / n, ch))
        chars.sort()
        res[label] = gaps_from_positions(chars, tk)
    report(res, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
