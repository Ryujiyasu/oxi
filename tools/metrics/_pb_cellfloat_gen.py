# -*- coding: utf-8 -*-
"""A floating drawing anchored INSIDE a table cell: does the row grow for it?

`creative__13152ea1`'s floating table holds a `wp:anchor` (wrapNone, positioned
relative to the paragraph, 256.9 x 120.75pt) in a 16pt exact-line cell paragraph
under `trHeight` 2604 (130.2pt). Word's row is 130.8 = the floor + border; Oxi's
is 147.75 = floor-less 16 + 120.75 + 11 -- S331b forwards every in-cell anchor as
flow height, and S1053 exempts only the PAGE-relative ones.

The candidate discriminator is the anchor's `layoutInCell` attribute (Word's
"Layout in table cell"): with it ON Word is known to grow the cell around the
object; with it OFF the object is laid out against the page and the cell keeps
its own height. The arms sweep that flag x wrap type x a row floor that is
SMALLER than the object, so growth is visible, and read the row as the span
[基準 -> 次] minus the same span of a no-object control.

    python _pb_cellfloat_gen.py gen
    python _pb_cellfloat_gen.py pdf      # Word truth (COM Info6)
    python _pb_cellfloat_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_cellfloat")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import NS  # noqa: E402
from _pb_exactimg_gen import CT, RELS, DRELS, DNS, png_bytes  # noqa: E402

MINCHO = "ＭＳ 明朝"
EMU = 12700
OBJ_H = 100.0            # object height in points; the small floor is 40pt

# (label, layoutInCell, wrap element, trHeight twips or None, object?)
ARMS = [
    ("control_floor40", 1, "wrapNone", 800, False),
    ("lic1_none_floor40", 1, "wrapNone", 800, True),
    ("lic0_none_floor40", 0, "wrapNone", 800, True),
    ("lic1_square_floor40", 1, "wrapSquare", 800, True),
    ("lic0_square_floor40", 0, "wrapSquare", 800, True),
    ("lic1_tb_floor40", 1, "wrapTopAndBottom", 800, True),
    ("lic0_tb_floor40", 0, "wrapTopAndBottom", 800, True),
    ("control_nofloor", 1, "wrapNone", None, False),
    ("lic1_none_nofloor", 1, "wrapNone", None, True),
    ("lic0_none_nofloor", 0, "wrapNone", None, True),
    # the witness's own shape: floor 130.2 larger than the 100pt object
    ("control_floor2604", 1, "wrapNone", 2604, False),
    ("lic1_none_floor2604", 1, "wrapNone", 2604, True),
    ("lic0_none_floor2604", 0, "wrapNone", 2604, True),
]


def docx(label):
    return os.path.join(OUT, "cellfloat_%s.docx" % label)


def anchor(lic, wrap):
    cx, cy = int(150 * EMU), int(OBJ_H * EMU)
    wrap_el = "<wp:%s/>" % wrap if wrap != "wrapSquare" else '<wp:wrapSquare wrapText="bothSides"/>'
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" '
            'simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" '
            'layoutInCell="%d" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>30480</wp:posOffset></wp:positionV>'
            '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '%s<wp:docPr id="1" name="pic1"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="pic1"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
            % (lic, cx, cy, wrap_el, cx, cy))


def para(text, extra_run=""):
    run = ('<w:r><w:t>%s</w:t></w:r>' % text) if text else ""
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/>'
            '</w:pPr>%s%s</w:p>' % (run, extra_run))


def table(inner, floor):
    trpr = "" if floor is None else '<w:trPr><w:trHeight w:val="%d"/></w:trPr>' % floor
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>'
            '<w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders>'
            '</w:tblPr><w:tblGrid><w:gridCol w:w="5394"/></w:tblGrid>'
            '<w:tr>%s<w:tc><w:tcPr><w:tcW w:w="5394" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>'
            % (trpr, inner))


def gen():
    os.makedirs(OUT, exist_ok=True)
    png = png_bytes()
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                '<w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              '<w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="22"/>'
              "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
              '<w:jc w:val="both"/></w:pPr></w:style></w:styles>'
              % (MINCHO, MINCHO, MINCHO))
    extra = " ".join(a for a in DNS.split(" ") if a.split("=")[0] + "=" not in NS)
    for label, lic, wrap, floor, has_obj in ARMS:
        cell = para("画像", anchor(lic, wrap) if has_obj else "") + para("説明")
        body = para("基準") + table(cell, floor) + para("次") + para("末尾")
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + " " + extra
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
                 '<w:docGrid w:type="linesAndChars" w:linePitch="375" w:charSpace="194"/>'
                 "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", DRELS)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/media/image1.png", png)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def control_of(label):
    return "control_" + label.split("_")[-1]


def report(spans, who, extra=None):
    print("== %s ==" % who)
    print("%-22s %-4s %-17s %-7s %-9s %-9s %s" % ("arm", "lic", "wrap", "floor", "基準->次", "-control", "notes"))
    for label, lic, wrap, floor, has_obj in ARMS:
        sp = spans.get(label)
        ctl = spans.get(control_of(label))
        d = None if (sp is None or ctl is None or not has_obj) else sp - ctl
        print("%-22s %-4d %-17s %-7s %-9s %-9s %s"
              % (label, lic, wrap, "-" if floor is None else "%.1f" % (floor / 20.0),
                 "-" if sp is None else "%.2f" % sp, "-" if d is None else "%+.2f" % d,
                 (extra or {}).get(label, "")))


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
    spans, extra = {}, {}
    for label, _, _, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "cellfloat_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "cf"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        by_y, borders = {}, []
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and (e.get("text") or "").strip():
                    by_y.setdefault(round(e["y"], 2), []).append((e["x"], e["text"]))
                elif e["type"] == "border" and e.get("w", 0) > 100:
                    borders.append(e["y"])
        y = {}
        for yy, frags in sorted(by_y.items()):
            t = "".join(t for _, t in sorted(frags)).strip()
            for key in ("基準", "次"):
                if t.startswith(key):
                    y.setdefault(key, yy)
        if "基準" in y and "次" in y:
            spans[label] = y["次"] - y["基準"]
        if len(borders) >= 2:
            extra[label] = "table h=%.2f" % (max(borders) - min(borders))
    report(spans, "OXI " + (envs or "(default)"), extra)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
