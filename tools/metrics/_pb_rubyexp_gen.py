# -*- coding: utf-8 -*-
"""Ruby line expansion: what does Word add for a ruby line, by face, base size,
hps and hpsRaise?

S1312 put the expansion on every ruby-bearing line; the value itself
(`ruby::ruby_expansion_pt`, calibrated on ＭＳ 明朝 10.5pt fixtures V1..V13) came
out 4.18 where Word gave 3.75 on `correspondence__04a3e3e1`'s HG丸ｺﾞｼｯｸ 10pt /
hps 5 / raise 9 lines. This sweeps the inputs so the formula can be checked
across faces and sizes rather than at one point.

Each arm stacks THREE identical one-line ruby paragraphs between 基準 and 次;
the per-line expansion is (span - the arm's no-ruby control span) / 3, so the
0.75pt Info6 quantisation is spent once over three lines.

    python _pb_rubyexp_gen.py gen
    python _pb_rubyexp_gen.py pdf      # Word truth (COM Info6)
    python _pb_rubyexp_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_rubyexp")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

FACES = {"hgmaru": "HG丸ｺﾞｼｯｸM-PRO", "msmincho": "ＭＳ 明朝", "msgothic": "ＭＳ ゴシック", "yumincho": "游明朝"}
N = 3

# (label, face key, base half-points, hps half-points or None (=base/2), hpsRaise half-points or None)
ARMS = []
for fk in ("hgmaru", "msmincho", "msgothic", "yumincho"):
    for base in (18, 20, 21, 24, 28):
        ARMS.append(("%s_b%d_ctl" % (fk, base), fk, base, None, None, False))
        ARMS.append(("%s_b%d_hpsHalf_rDef" % (fk, base), fk, base, base // 2, None, True))
        ARMS.append(("%s_b%d_hpsHalf_r18" % (fk, base), fk, base, base // 2, 18, True))
# the witness's exact combination and a small/large-raise pair on it
ARMS.append(("hgmaru_b20_hps10_r18", "hgmaru", 20, 10, 18, True))
ARMS.append(("hgmaru_b20_hps10_r12", "hgmaru", 20, 10, 12, True))
ARMS.append(("hgmaru_b20_hps10_r24", "hgmaru", 20, 10, 24, True))
ARMS.append(("hgmaru_b20_hps8_r18", "hgmaru", 20, 8, 18, True))
ARMS.append(("hgmaru_b20_hps14_r18", "hgmaru", 20, 14, 18, True))


def docx(label):
    return os.path.join(OUT, "rubyexp_%s.docx" % label)


def fonts(face):
    return '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>' % (face, face, face)


def para(face, base, hps, raise_, with_ruby):
    rpr = fonts(face) + '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (base, base)
    if not with_ruby:
        run = '<w:r><w:rPr>%s</w:rPr><w:t>年金相談会です</w:t></w:r>' % rpr
    else:
        rp = '<w:rubyAlign w:val="distributeSpace"/><w:hps w:val="%d"/>' % hps
        if raise_ is not None:
            rp += '<w:hpsRaise w:val="%d"/>' % raise_
        rp += '<w:hpsBaseText w:val="%d"/><w:lid w:val="ja-JP"/>' % base
        run = ('<w:r><w:rPr>%s</w:rPr><w:ruby><w:rubyPr>%s</w:rubyPr><w:rt><w:r><w:rPr>%s<w:sz w:val="%d"/></w:rPr>'
               '<w:t>ねんきん</w:t></w:r></w:rt><w:rubyBase><w:r><w:rPr>%s</w:rPr><w:t>年金</w:t></w:r></w:rubyBase></w:ruby></w:r>'
               '<w:r><w:rPr>%s</w:rPr><w:t>相談会です</w:t></w:r>' % (rpr, rp, fonts(face), hps, rpr, rpr))
    return ('<w:p><w:pPr><w:widowControl/><w:snapToGrid w:val="0"/><w:rPr>%s</w:rPr></w:pPr>%s</w:p>' % (rpr, run))


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
                '<w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
              '<w:jc w:val="both"/></w:pPr></w:style></w:styles>')
    for label, fk, base, hps, raise_, with_ruby in ARMS:
        body = marker("基準") + para(FACES[fk], base, hps, raise_, with_ruby) * N + marker("次") + marker("末尾")
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
    print("%-26s %-14s %-5s %-5s %-6s %s" % ("arm", "face", "base", "hps", "raise", "expansion / line"))
    for label, fk, base, hps, raise_, with_ruby in ARMS:
        if not with_ruby:
            continue
        sp = spans.get(label); ctl = spans.get("%s_b%d_ctl" % (fk, base))
        d = None if (sp is None or ctl is None) else (sp - ctl) / N
        print("%-26s %-14s %-5.1f %-5.1f %-6s %s" % (label, FACES[fk], base / 2.0, hps / 2.0,
                                                    "def" if raise_ is None else "%.1f" % (raise_ / 2.0),
                                                    "-" if d is None else "%+.3f" % d))


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    spans = {}
    try:
        for label, _, _, _, _, _ in ARMS:
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
    for label, _, _, _, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "rubyexp_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "re"),
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
