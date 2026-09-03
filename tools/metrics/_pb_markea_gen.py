# -*- coding: utf-8 -*-
"""An empty paragraph whose ¶ MARK asks for the East Asian THEME — which face
does Word actually give the line?

`f16f228a` has two such paragraphs inside a table:

    <w:pPr><w:snapToGrid w:val="0"/>
      <w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/><w:sz w:val="18"/></w:rPr>
    </w:pPr>

The theme's Jpan face is 游明朝, so "resolve the theme" gives 15.00pt at 9pt,
while the ASCII font it inherits (ＭＳ 明朝, from the Normal style) gives 11.67.
Word's own PDF puts the following text where the 11.67 reading predicts -- the
ASCII rule (S583/S707) wins even though the mark names the EA theme outright.
That could not be seen before: whenever the theme resolved to ＭＳ 明朝 as well,
the two readings gave the same number.

Each arm is [marker] [the empty paragraph] [marker]; the arm's height is the
marker-to-marker distance minus the `control` arm's.

    python _pb_markea_gen.py gen
    python _pb_markea_gen.py pdf      # Word truth
    python _pb_markea_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_markea")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402
from _pb_themeea_gen import theme_xml  # noqa: E402

MINCHO = "ＭＳ 明朝"
SZ = 18          # half-points = 9pt, as in f16f228a

# (label, the Normal style's ascii face, the mark's rPr rFonts, grid?)
ARMS = [
    ("control", MINCHO, None, False),
    ("mark_eatheme_ascii_cjk", MINCHO, '<w:rFonts w:eastAsiaTheme="minorEastAsia"/>', False),
    ("mark_eatheme_ascii_latin", "Century", '<w:rFonts w:eastAsiaTheme="minorEastAsia"/>', False),
    ("mark_ea_literal", MINCHO, '<w:rFonts w:eastAsia="%s"/>' % MINCHO, False),
    ("mark_no_rfonts", MINCHO, None, False),
    # The witness sits in a table with snapToGrid=0; a typed grid takes a
    # different branch in Oxi, so sweep that too.
    ("grid_mark_eatheme_ascii_cjk", MINCHO,
     '<w:rFonts w:eastAsiaTheme="minorEastAsia"/>', True),
    ("grid_control", MINCHO, None, True),
]


def docx(label):
    return os.path.join(OUT, "markea_%s.docx" % label)


def marker(ascii_face):
    return ('<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
            '<w:sz w:val="%d"/></w:rPr><w:t>基準</w:t></w:r></w:p>'
            % (ascii_face, ascii_face, SZ))


def gen():
    os.makedirs(OUT, exist_ok=True)
    ct = CT.replace(
        "</Types>",
        '<Override PartName="/word/theme/theme1.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rel = ('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
           'officeDocument/2006/relationships/%s" Target="%s"/>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
             'relationships">'
             + rel % (1, "styles", "styles.xml")
             + rel % (2, "theme", "theme/theme1.xml")
             + rel % (3, "settings", "settings.xml") + "</Relationships>")
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                '<w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
    theme = theme_xml("", {"Jpan": "游明朝"})
    for label, ascii_face, mark_fonts, grid in ARMS:
        # docDefaults names ＭＳ 明朝 for eastAsia outright, exactly as the witness
        # does, so only a mark that asks for the theme by name can reach it.
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  "<w:docDefaults><w:rPrDefault><w:rPr>"
                  '<w:rFonts w:ascii="Century" w:eastAsia="%s" w:hAnsi="Century"/>'
                  '<w:sz w:val="%d"/></w:rPr></w:rPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
                  '<w:name w:val="Normal"/><w:rPr>'
                  '<w:rFonts w:ascii="%s" w:hAnsi="%s"/></w:rPr></w:style></w:styles>'
                  % (MINCHO, SZ, ascii_face, ascii_face))
        mid = ""
        if mark_fonts is not None:
            mid = ('<w:p><w:pPr><w:snapToGrid w:val="0"/><w:rPr>%s'
                   '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:pPr></w:p>'
                   % (mark_fonts, SZ, SZ))
        elif label.endswith("no_rfonts"):
            mid = ('<w:p><w:pPr><w:snapToGrid w:val="0"/><w:rPr>'
                   '<w:sz w:val="%d"/></w:rPr></w:pPr></w:p>' % SZ)
        grid_el = ('<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else "")
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + marker(ascii_face) + mid + marker(ascii_face)
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
               + grid_el + "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/theme/theme1.xml", theme)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def report(rows):
    base = {}
    print("%-30s %-9s %-24s %-6s %-8s %s"
          % ("arm", "ascii", "mark rFonts", "grid", "height", "gap"))
    for label, ys in rows:
        marks = [y for y, t in ys if "基準" in t]
        gap = (max(marks) - min(marks)) if len(marks) >= 2 else None
        key = "grid" if label.startswith("grid") else "plain"
        if label.endswith("control"):
            base[key] = gap
        h = "" if (gap is None or key not in base) else "%.2f" % (gap - base[key])
        a = dict((l, (af, mf, g)) for l, af, mf, g in ARMS)[label]
        mf = (a[1] or "(none)").replace("<w:rFonts ", "").replace("/>", "")
        print("%-30s %-9s %-24s %-6s %-8s %s"
              % (label, a[0], mf[:24], a[2], h, "-" if gap is None else "%.2f" % gap))


def pdf():
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    rows = []
    try:
        for label, _, _, _ in ARMS:
            src, dst = docx(label), docx(label).replace(".docx", ".pdf")
            d = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            try:
                d.ExportAsFixedFormat(dst, 17)
            finally:
                d.Close(False)
            ys = []
            for b in fitz.open(dst)[0].get_text("dict")["blocks"]:
                for ln in b.get("lines", []):
                    t = "".join(s["text"] for s in ln["spans"])
                    if t.strip():
                        ys.append((min(s["bbox"][1] for s in ln["spans"]), t))
            rows.append((label, ys))
    finally:
        app.Quit()
    print("== WORD ==")
    report(rows)


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    rows = []
    for label, _, _, _ in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "markea_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "me"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        by_y = {}
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and (e.get("text") or "").strip():
                    by_y.setdefault(round(e["y"], 2), []).append((e["x"], e["text"]))
        rows.append((label, [(y, "".join(t for _, t in sorted(v))) for y, v in sorted(by_y.items())]))
    print("== OXI %s ==" % (envs or "(default)"))
    report(rows)


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
