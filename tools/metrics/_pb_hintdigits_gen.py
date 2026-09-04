# -*- coding: utf-8 -*-
"""A line of DIGITS under `w:hint="eastAsia"` -- whose natural height is it?

`correspondence__04a3e3e1`'s timetable has date rows: one 9pt run of digits in
HG丸ｺﾞｼｯｸM-PRO with `<w:rFonts ... w:hint="eastAsia"/>`, `snapToGrid=0`,
`trHeight` 63tw (no floor). Word advances those rows 12.00 (the 0.75pt
quantisation of 11.67 = 9 x 83/64, the CJK natural) where Oxi gives 10.31
(9 x 1.146). Five such rows on the page = -8.5pt.

The arms separate the three candidate inputs: the FACE (HG丸ｺﾞｼｯｸ / ＭＳ ゴシック /
ＭＳ 明朝 / Century), the HINT (eastAsia or absent) and the CONTENT (digits /
CJK / Latin letters). Each arm stacks THREE identical single-line paragraphs
between 基準 and 次 so the per-line height is the span difference / 3 -- the
0.75pt Info6 quantisation is spent once over three lines.

    python _pb_hintdigits_gen.py gen
    python _pb_hintdigits_gen.py pdf      # Word truth (COM Info6)
    python _pb_hintdigits_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_hintdigits")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

SZ = 18   # 9pt
N = 3     # stacked test lines per arm
FACES = {"hgmaru": "HG丸ｺﾞｼｯｸM-PRO", "msgothic": "ＭＳ ゴシック", "msmincho": "ＭＳ 明朝", "century": "Century"}
CONTENT = {"digits": "23", "cjk": "月", "latin": "ab"}

# (label, face key, hint?, content key)
ARMS = [("control", "hgmaru", False, None)]
for fk in ("hgmaru", "msgothic", "msmincho", "century"):
    for ck in ("digits", "cjk", "latin"):
        for hint in (True, False):
            if ck == "cjk" and not hint:
                continue
            ARMS.append(("%s_%s_%s" % (fk, ck, "hint" if hint else "nohint"), fk, hint, ck))
# Table arms: the witness's date rows are CELLS (trHeight 63tw, tblCellMar 99/99,
# snapToGrid=0). Four stacked one-cell rows, digits 21..24, per-row pitch read
# from consecutive rows so no control is needed.
TBL_ARMS = [("tbl_hgmaru_digits_hint", "hgmaru", True), ("tbl_hgmaru_digits_nohint", "hgmaru", False),
            ("tbl_msgothic_digits_hint", "msgothic", True), ("tbl_century_digits_hint", "century", True),
            ("tbl_hgmaru_cjk_hint", "hgmaru", None),
            # the witness's table FLOATS (tblpPr vertAnchor=text); a float takes
            # Oxi's pre-pass row estimate, so sweep that too
            ("ftbl_hgmaru_digits_hint", "hgmaru", True), ("ftbl_century_digits_hint", "century", True)]
ROWS = 4
# (label, face, hint, adjustLineHeightInTable?, snapToGrid?) -- the witness has
# `<w:adjustLineHeightInTable/>` in its compat block and `snapToGrid=0` on the
# date-row paragraphs; Oxi's ALIT arm drops the 83/64 natural for those cells.
TBL2_ARMS = [("alit1_snap0_hgmaru", "hgmaru", True, True, False),
             ("alit1_snap1_hgmaru", "hgmaru", True, True, True),
             ("alit0_snap1_hgmaru", "hgmaru", True, False, True),
             ("alit1_snap0_century", "century", True, True, False),
             ("alit1_snap1_century", "century", True, True, True),
             ("alit0_snap1_century", "century", True, False, True),
             ("alit1_snap0_msgothic", "msgothic", True, True, False),
             ("alit1_snap1_msgothic", "msgothic", True, True, True)]
TBLP = ('<w:tblpPr w:leftFromText="142" w:rightFromText="142" w:vertAnchor="text" '
        'w:horzAnchor="margin" w:tblpY="1"/>')


def docx(label):
    return os.path.join(OUT, "hintdigits_%s.docx" % label)


def para(text, face, hint, snap=False):
    h = ' w:hint="eastAsia"' if hint else ""
    fonts = '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:cs="ＭＳ Ｐゴシック"%s/>' % (face, face, face, h)
    return ('<w:p><w:pPr><w:widowControl/><w:adjustRightInd w:val="0"/>'
            + ("" if snap else '<w:snapToGrid w:val="0"/>') +
            '<w:jc w:val="center"/><w:rPr>%s<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:pPr>'
            '<w:r><w:rPr>%s<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr><w:t>%s</w:t></w:r></w:p>'
            % (fonts, SZ, SZ, fonts, SZ, SZ, text))


def table_rows(face, hint, floating=False, snap=False):
    out = ('<w:tbl><w:tblPr>' + (TBLP if floating else "") + '<w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
           '<w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>'
           '<w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>'
           '<w:insideH w:val="single" w:sz="4"/></w:tblBorders>'
           '<w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar>'
           '</w:tblPr><w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid>')
    for i in range(ROWS):
        txt = "月" if hint is None else str(21 + i)
        out += ('<w:tr><w:trPr><w:trHeight w:val="63"/></w:trPr><w:tc><w:tcPr>'
                '<w:tcW w:w="1200" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr>'
                % para(txt, FACES[face], bool(hint), snap=snap))
    return out + "</w:tbl>"


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
    # docDefaults as the witness: Century / ＭＳ 明朝, no size (10.5), pPrDefault empty
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
              '<w:jc w:val="both"/></w:pPr></w:style></w:styles>')
    alit = {l: a for l, _f, _h, a, _s in TBL2_ARMS}
    snap = {l: s_ for l, _f, _h, _a, s_ in TBL2_ARMS}
    for label, fk, hint, ck in (ARMS + [(l, f, h, "tbl") for l, f, h in TBL_ARMS]
                                + [(l, f, h, "tbl") for l, f, h, _a, _s in TBL2_ARMS]):
        body = marker("基準")
        if ck == "tbl":
            body += table_rows(fk, hint, floating=label.startswith("ftbl_"), snap=snap.get(label, False))
        elif ck is not None:
            for _ in range(N):
                body += para(CONTENT[ck], FACES[fk], hint)
        body += marker("次") + marker("末尾")
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
                 '<w:docGrid w:type="lines" w:linePitch="344"/>'
                 "</w:sectPr></w:body></w:document>")
        st = settings
        if alit.get(label):
            st = st.replace("<w:compat>", "<w:compat><w:adjustLineHeightInTable/>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", st)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS) + len(TBL_ARMS) + len(TBL2_ARMS), OUT))


def report(spans, who, rowpitch=None):
    print("== %s ==" % who)
    for label, fk, hint in TBL_ARMS + [(l, f, h) for l, f, h, _a, _s in TBL2_ARMS]:
        rp = (rowpitch or {}).get(label)
        print("%-26s %-16s %-7s %-7s %-9s %s" % (label, FACES[fk], "eastAsia" if hint else ("cjk" if hint is None else "-"),
              "cell", "-" if rp is None else "%.3f" % rp, "-" if rp is None else "%.3f" % (rp / 9.0)))
    ctl = spans.get("control")
    print("%-26s %-16s %-7s %-7s %-9s %s" % ("arm", "face", "hint", "text", "per line", "x 9pt"))
    for label, fk, hint, ck in ARMS:
        sp = spans.get(label)
        if ck is None or sp is None or ctl is None:
            continue
        per = (sp - ctl) / N
        print("%-26s %-16s %-7s %-7s %-9.3f %.3f" % (label, FACES[fk], "eastAsia" if hint else "-", ck, per, per / 9.0))


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    spans, rowpitch = {}, {}
    try:
        for label, _, _, _ in (ARMS + [(l, f, h, "tbl") for l, f, h in TBL_ARMS]
                               + [(l, f, h, "tbl") for l, f, h, _a, _s in TBL2_ARMS]):
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                ys = {}
                rows = []
                for i in range(1, d.Paragraphs.Count + 1):
                    p = d.Paragraphs(i)
                    st = d.Range(p.Range.Start, p.Range.Start)
                    t = (p.Range.Text or "").rstrip("\r\x07")
                    ys.setdefault(t, float(st.Information(6)))
                    if t in ("21", "22", "23", "24", "月"):   # the table rows
                        rows.append(float(st.Information(6)))
                spans[label] = ys["次"] - ys["基準"]
                if len(rows) >= 2:
                    rowpitch[label] = (rows[-1] - rows[0]) / (len(rows) - 1)
            finally:
                d.Close(False)
    finally:
        app.Quit()
    report(spans, "WORD (Info6, collapsed starts)", rowpitch)


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    spans, rowpitch = {}, {}
    for label, _, _, _ in (ARMS + [(l, f, h, "tbl") for l, f, h in TBL_ARMS]
                           + [(l, f, h, "tbl") for l, f, h, _a, _s in TBL2_ARMS]):
        dump = os.path.join(tempfile.gettempdir(), "hintdigits_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "hd"),
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
        rows = sorted(yy for yy, frags in by_y.items()
                      if "".join(t for _, t in frags).strip() in ("21", "22", "23", "24", "月"))
        if len(rows) >= 2:
            rowpitch[label] = (rows[-1] - rows[0]) / (len(rows) - 1)
    report(spans, "OXI " + (envs or "(default)"), rowpitch)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
