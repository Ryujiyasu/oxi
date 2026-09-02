# -*- coding: utf-8 -*-
"""How much room does an INLINE text box need before Word moves it to the next
page -- and does Oxi demand the same?

`legal__02f84965dccfe4db` p5: Word keeps a 455 x 103.35pt inline shape
(`wps:wsp` inside `wp:inline`, its own paragraphs living in `w:txbxContent`) on
page 4 with ~11pt to spare; Oxi pushes it to page 5 even though it has 113.05pt
free -- 6 paragraphs then fall off the page and the rest of the document runs
+1. So Oxi demands MORE room for the box than Word does, and this finds how
much more by sweeping the space left above it.

The shape XML is a FAITHFUL SLICE of the real document, not a hand-written box:
a minimal shape makes Word fall back to a degraded layout path and the threshold
it reports is then not the one the real document meets
([[probe_minimal_docx_degraded]]).

Geometry: A4 portrait, 70.9pt margins -> content height 700.15pt. Each filler
paragraph is pinned to `exact` 12.0pt, so N fillers put the cursor at exactly
70.9 + 12N and the room left for the box is 700.15 - 12N. Sweeping N one step
at a time resolves the threshold to 12pt; the arms bracket 103.35 generously.

    python _pb_inlinebox_gen.py gen
    python _pb_inlinebox_gen.py measure     # Word -> PDF
    python _pb_inlinebox_gen.py read        # Word truth
    python _pb_inlinebox_gen.py oxi         # Oxi
"""
import glob
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUTDIR = os.path.join(REPO, "pipeline_data", "_pb_inlinebox")
SRC = os.path.join(REPO, "pipeline_data", "docx_corpus", "ja", "legal",
                   "02f84965dccfe4db.docx")

BOX_H = 103.35          # cy 1312545 EMU
CONTENT_H = 841.95 - 2 * 70.9
FILL = 12.0
# room left above the box = CONTENT_H - FILL*N. Bracket BOX_H (103.35) widely:
# N=48 -> 124.15 room, N=53 -> 64.15 room.
NS = list(range(48, 54))


def slice_shape():
    """Pull the real AlternateContent block and the document's ns declaration."""
    xml = zipfile.ZipFile(SRC).read("word/document.xml").decode("utf-8")
    off = xml.index('cy="1312545"')
    s = xml.rfind("<mc:AlternateContent", 0, off)
    i, depth = s, 0
    pat = re.compile(r"</?mc:AlternateContent\b")
    while True:
        m = pat.search(xml, i)
        if not m:
            break
        depth += -1 if xml[m.start():m.start() + 2] == "</" else 1
        i = m.end()
        if depth == 0:
            # `\b` stops before the tag's own `>`; take it too, or the slice is
            # ill-formed and the renderer refuses the whole document.
            i = xml.index(">", i) + 1
            break
    return xml[s:i], re.search(r"<w:document[^>]*>", xml).group(0)


CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
DOCRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""
SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:compatSetting w:name="compatibilityMode"
 w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"
 w:eastAsia="ＭＳ ゴシック"/><w:sz w:val="18"/><w:szCs w:val="18"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"'
        ' w:header="737" w:footer="340" w:gutter="0"/></w:sectPr>')


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    shape, decl = slice_shape()
    for n, tail in [(n, False) for n in NS] + [(n, True) for n in NS]:
        fillers = "".join(
            f'<w:p><w:r><w:t>FILL{i:03d}</w:t></w:r></w:p>' for i in range(n))
        # The fillers are pinned to `exact` 12pt so the sweep step is exact, but
        # the BOX paragraph must NOT inherit that: an exact line height clamps
        # the line to 12pt, Word then draws the shape overlapping the text above
        # it and reserves nothing, and the arm measures the clamp instead of the
        # break. (First cut of this probe did exactly that -- Word answered
        # "1 page" for every arm with the box drawn ABOVE the last filler.)
        auto = '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>'
        boxp = f"<w:p>{auto}<w:r>{shape}</w:r></w:p>"
        after = f'<w:p>{auto}<w:r><w:t>AFTERBOX</w:t></w:r></w:p>'
        # `t` arms: the shape run is followed, IN THE SAME PARAGRAPH, by a run
        # holding `<w:br w:type="page"/>` -- the shape of the real document.
        # Deleting that one run from legal__02f84965 moves the box from page 5
        # to page 4 and the document from 11 pages to 10, so this is the
        # discriminator. A trailing EMPTY run (no break) was tried first and
        # changed nothing in either engine: it is the BREAK that matters, not
        # the extra run.
        if tail:
            boxp = (f"<w:p>{auto}<w:r>{shape}</w:r>"
                    f'<w:r><w:rPr><w:sz w:val="18"/></w:rPr>'
                    f'<w:br w:type="page"/></w:r></w:p>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               + decl + "<w:body>" + fillers + boxp + after + SECT + "</w:body></w:document>")
        path = os.path.join(OUTDIR, ("t" if tail else "n") + f"{n}.docx")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", DOCRELS)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", SETTINGS)
        print(f"  {'t' if tail else 'n'}{n:>3} fillers -> room above the box = "
              f"{CONTENT_H - FILL * n:7.2f}pt  (box needs {BOX_H})")


def measure():
    import win32com.client as win32
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for path in sorted(glob.glob(os.path.join(OUTDIR, "*.docx"))):
            d = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            try:
                d.ExportAsFixedFormat(
                    OutputFileName=os.path.abspath(path[:-5] + ".pdf"), ExportFormat=17)
                print("measured", os.path.basename(path))
            finally:
                d.Close(False)
    finally:
        word.Quit()


def read():
    import fitz
    print(f"{'arm':<6} {'room':>8} {'pages':>5}  box page  (Word)")
    for path in sorted(glob.glob(os.path.join(OUTDIR, "*.pdf"))):
        arm = os.path.basename(path)[:-4]; n = int(arm[1:])
        doc = fitz.open(path)
        where = None
        for i, pg in enumerate(doc):
            # the box carries the shape's own text; AFTERBOX follows it
            if "宛先" in pg.get_text() and where is None:
                where = i + 1
        print(f"{arm:<6} {CONTENT_H - FILL*n:>8.2f} {len(doc):>5}  {where}")


def oxi():
    import json, subprocess, tempfile
    exe = os.environ.get("OXI_GDI_EXE") or os.path.join(
        REPO, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
    print(f"{'arm':<6} {'room':>8} {'pages':>5}  box page  (Oxi)")
    for path in sorted(glob.glob(os.path.join(OUTDIR, "*.docx"))):
        arm = os.path.basename(path)[:-5]; n = int(arm[1:])
        with tempfile.TemporaryDirectory() as t:
            dj = os.path.join(t, "l.json")
            r = subprocess.run([exe, os.path.abspath(path), os.path.join(t, "p"),
                                "--dump-layout=" + dj], capture_output=True, timeout=180)
            if r.returncode != 0 or not os.path.exists(dj):
                print(f"{arm:<6}  RENDER FAIL")
                continue
            dump = json.load(open(dj, encoding="utf-8"))
        where = None
        for i, pg in enumerate(dump["pages"], 1):
            if any(e.get("type") == "image" for e in pg["elements"]) and where is None:
                where = i
        print(f"{arm:<6} {CONTENT_H - FILL*n:>8.2f} {len(dump['pages']):>5}  {where}")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    {"gen": gen, "measure": measure, "read": read, "oxi": oxi}[cmd]()
