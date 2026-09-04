# -*- coding: utf-8 -*-
"""A BLANK line under a multiplier in a typed grid -- what does Word advance it by,
and does the paragraph AFTER it have a say?

`creative__13152ea1` (linesAndChars, pitch 375tw = 18.75pt, ＭＳ 明朝 12pt) opens
with four empty paragraphs at `w:line="276" w:lineRule="auto"` (x1.15) and then a
14pt `w:line="520" w:lineRule="exact"` heading. Word's own advances (COM):

    21.00 / 21.75 / 21.75 / 18.75

The S1306 law (`pitch x max(ceil(nat / pitch), mult)`) gives 21.56 for every one
of them; the first three ARE that within Info6's 0.75pt quantisation, the last
is one bare cell. Two readings fit: "the last blank of a run" and "the blank
before an EXACT paragraph". The arms below differ in exactly those two things,
and add a TEXT run before the same exact paragraph, because the S1306 sweep's
eight arms never had an exact paragraph after them -- if a text line drops the
same way, the rule is about the FOLLOWING paragraph, not about blanks.

Every arm is [基準] [N blanks or text lines] [次 = the paragraph under test]
[末尾]. Word reports each paragraph's Info6 (collapsed start, per CLAUDE.local.md);
Oxi cannot see a blank, so both sides are ALSO reported as the span 基準 -> 次
minus the same span in the arm's zero-blank control, which cancels the marker
line's own height and any text_y_off.

    python _pb_blankmult_gen.py gen
    python _pb_blankmult_gen.py pdf      # Word truth (COM Info6)
    python _pb_blankmult_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_blankmult")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

MINCHO = "ＭＳ 明朝"
LC = '<w:docGrid w:type="linesAndChars" w:linePitch="375" w:charSpace="194"/>'
LN = '<w:docGrid w:type="lines" w:linePitch="375"/>'
TEXT = "本文の行です。"

# (label, grid, n, w:line of the run, run is TEXT?, (rule, w:line) of 次)
ARMS = [
    ("w_b4_x115_exact520", LC, 4, 276, False, ("exact", 520)),    # the witness
    ("b0_exact520", LC, 0, 276, False, ("exact", 520)),           # control for the above
    ("b4_x115_auto276", LC, 4, 276, False, ("auto", 276)),        # next = same rule as the blanks
    ("b0_auto276", LC, 0, 276, False, ("auto", 276)),             # control
    ("b4_x115_auto240", LC, 4, 276, False, ("auto", 240)),
    ("b0_auto240", LC, 0, 276, False, ("auto", 240)),             # control
    ("b4_x115_exact240", LC, 4, 276, False, ("exact", 240)),
    ("b0_exact240", LC, 0, 276, False, ("exact", 240)),           # control
    ("b4_x115_atleast400", LC, 4, 276, False, ("atLeast", 400)),
    ("b0_atleast400", LC, 0, 276, False, ("atLeast", 400)),       # control
    ("t4_x115_exact520", LC, 4, 276, True, ("exact", 520)),       # TEXT lines, then exact
    ("b1_x115_exact520", LC, 1, 276, False, ("exact", 520)),
    ("b4_x150_exact520", LC, 4, 360, False, ("exact", 520)),
    ("b4_x100_exact520", LC, 4, 240, False, ("exact", 520)),      # no multiplier
    ("lines_b4_x115_exact520", LN, 4, 276, False, ("exact", 520)),
    ("lines_b0_exact520", LN, 0, 276, False, ("exact", 520)),     # control
]


def docx(label):
    return os.path.join(OUT, "blankmult_%s.docx" % label)


def para(text, rule, line, sz=None):
    rpr = "" if sz is None else '<w:rPr><w:sz w:val="%d"/></w:rPr>' % sz
    run = "" if not text else ('<w:r>%s<w:t>%s</w:t></w:r>' % (rpr, text))
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="%s"/>'
            '<w:ind w:firstLineChars="100" w:firstLine="241"/></w:pPr>%s</w:p>'
            % (line, rule, run))


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
    # docDefaults as the witness has them: ＭＳ 明朝 12pt, kern 2, Normal = widowControl 0 + jc both.
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              '<w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="22"/>'
              "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
              '<w:jc w:val="both"/></w:pPr></w:style></w:styles>'
              % (MINCHO, MINCHO, MINCHO))
    for label, grid, n, line, is_text, (nrule, nline) in ARMS:
        body = para("基準", "auto", 240)
        for _ in range(n):
            body += para(TEXT if is_text else "", "auto", line)
        # 次 is the witness's heading: 14pt under the rule under test.
        body += para("次", nrule, nline, sz=28)
        body += para("末尾", "auto", 240)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
               + grid + "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def control_of(label):
    if label.startswith("lines_"):
        return "lines_b0_exact520"
    tail = label.split("_")[-1]          # exact520 / auto276 / ...
    return "b0_" + tail


def report(spans, who, advances=None):
    print("== %s ==" % who)
    print("%-24s %-6s %-4s %-5s %-5s %-9s %-10s %-10s %s"
          % ("arm", "grid", "n", "run", "text", "next", "span", "-control", "per line"))
    for label, grid, n, line, is_text, (nrule, nline) in ARMS:
        sp = spans.get(label)
        ctl = spans.get(control_of(label))
        blanks = None if (sp is None or ctl is None) else sp - ctl
        per = "" if (blanks is None or n == 0) else "%.2f" % (blanks / n)
        print("%-24s %-6s %-4d x%-4.2f %-5s %-9s %-10s %-10s %s"
              % (label, "lines" if grid == LN else "l&c", n, line / 240.0,
                 "text" if is_text else "blank", "%s%d" % (nrule, nline),
                 "-" if sp is None else "%.2f" % sp,
                 "-" if blanks is None else "%.2f" % blanks, per))
        if advances and label in advances:
            print("%-24s     Word advances: %s" % ("", "  ".join("%.2f" % a for a in advances[label])))


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    spans, advances = {}, {}
    try:
        for label, _, _, _, _, _ in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                ys, texts = [], []
                for i in range(1, d.Paragraphs.Count + 1):
                    p = d.Paragraphs(i)
                    st = d.Range(p.Range.Start, p.Range.Start)
                    ys.append(float(st.Information(6)))
                    texts.append((p.Range.Text or "").rstrip("\r\x07"))
            finally:
                d.Close(False)
            advances[label] = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
            spans[label] = ys[texts.index("次")] - ys[texts.index("基準")]
    finally:
        app.Quit()
    report(spans, "WORD (Info6, collapsed starts)", advances)


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    spans = {}
    for label, _, _, _, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "blankmult_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "bm"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        # The dump may split a line into fragments: join them by y first.
        by_y = {}
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and (e.get("text") or "").strip():
                    by_y.setdefault(round(e["y"], 2), []).append((e["x"], e["text"]))
        y = {}
        for yy, frags in sorted(by_y.items()):
            t = "".join(t for _, t in sorted(frags))
            for key in ("基準", "次"):
                if key in t and t.strip() != "末尾":
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
