# -*- coding: utf-8 -*-
"""When does Word insert a BLANK page at a nextPage section break?

legal__0010437a (WA Supreme Court Rules) renders a cover page, then a
COMPLETELY empty physical page 2, then its numbered content from page 3 —
even though every sectPr is a plain nextPage break (no evenPage/oddPage, so
S732's parity rule does not fire) and nothing on page 2 is even a paragraph
mark. Oxi produces no such page, so its whole body sits one page early.

The document's distinguishing features are
  settings.xml : <w:evenAndOddHeaders/>
  sect0        : titlePg, pgNumType fmt=lowerRoman start=1
  sect1        : titlePg, pgNumType start=1        <- restarts at LOGICAL 1

Hypothesis: with different odd/even headers in force, a section that restarts
page numbering at an ODD number must begin on a page of matching parity, so
Word pads with a blank — the pgNumType-driven sibling of S732's
evenPage/oddPage rule.

Readout: the number of pages, and which page carries the section-2 marker
text. A blank page shows up as a page whose extracted text is empty.

Usage: python _pb_blankpage_gen.py gen | measure | read
"""
import os, sys, glob, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUTDIR = os.path.join(REPO, "pipeline_data", "_pb_blankpage")

# (id, even_odd_headers, sect0_pgnum, sect1_pgnum, sect0_titlepg, sect1_titlepg)
CASES = [
    # the l10 replica
    ("b1_eoh_restart1",   True,  'w:fmt="lowerRoman" w:start="1"', 'w:start="1"', True,  True),
    # is evenAndOddHeaders the switch?
    ("b2_noeoh_restart1", False, 'w:fmt="lowerRoman" w:start="1"', 'w:start="1"', True,  True),
    # is the RESTART the switch (drop sect1's pgNumType)?
    ("b3_eoh_norestart",  True,  'w:fmt="lowerRoman" w:start="1"', None,          True,  True),
    # restart at an EVEN number -> should need no pad if parity is the rule
    ("b4_eoh_restart2",   True,  'w:fmt="lowerRoman" w:start="1"', 'w:start="2"', True,  True),
    # is titlePg involved?
    ("b5_eoh_notitlepg",  True,  'w:fmt="lowerRoman" w:start="1"', 'w:start="1"', False, False),
    # plain control: no pgNumType anywhere
    ("b6_eoh_nopgnum",    True,  None,                            None,          True,  True),
    # 2-page cover -> section 2 would land on physical 3 (ODD).
    # restart at 2 (EVEN): parity-match rule pads, odd-restart rule does not.
    ("b7_2pg_restart2",   True,  'w:fmt="lowerRoman" w:start="1"', 'w:start="2"', True,  True),
    # restart at 1 (ODD) onto physical 3: both rules predict no pad.
    ("b8_2pg_restart1",   True,  'w:fmt="lowerRoman" w:start="1"', 'w:start="1"', True,  True),
]

# ids whose cover section spans two pages (a page break inside the cover)
TWO_PAGE_COVER = {"b7_2pg_restart2", "b8_2pg_restart1"}

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

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>"""


def settings_xml(eoh):
    flag = "<w:evenAndOddHeaders/>" if eoh else ""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
{flag}<w:compat><w:compatSetting w:name="compatibilityMode"
 w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>"""


PGSZ = ('<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/>')


def sect(pgnum, titlepg, inner=True):
    """A sectPr; `inner` wraps it in the trailing paragraph of the section."""
    body = PGSZ
    if pgnum:
        body += f'<w:pgNumType {pgnum}/>'
    if titlepg:
        body += "<w:titlePg/>"
    s = f"<w:sectPr>{body}</w:sectPr>"
    return f"<w:p><w:pPr>{s}</w:pPr></w:p>" if inner else s


def para(t):
    return f'<w:p><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for (cid, eoh, pg0, pg1, tp0, tp1) in CASES:
        cover = para("COVER PAGE")
        if cid in TWO_PAGE_COVER:
            cover += ('<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                      + para("COVER PAGE 2"))
        body = (cover + sect(pg0, tp0)
                + para("SECTION2 FIRST LINE") + para("more text")
                + sect(pg1, tp1, inner=False))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
               f'<w:body>{body}</w:body></w:document>')
        with zipfile.ZipFile(os.path.join(OUTDIR, cid + ".docx"), "w",
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", DOCRELS)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings_xml(eoh))
        print("gen", cid)


def measure():
    import win32com.client as win32
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        for path in sorted(glob.glob(os.path.join(OUTDIR, "*.docx"))):
            d = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            try:
                d.ExportAsFixedFormat(OutputFileName=os.path.abspath(path[:-5] + ".pdf"),
                                      ExportFormat=17)
                print("measured", os.path.basename(path))
            finally:
                d.Close(False)
    finally:
        word.Quit()


def read():
    import fitz
    print(f"{'case':<22} {'pages':>5}  {'sect2 on':>8}  blanks")
    for path in sorted(glob.glob(os.path.join(OUTDIR, "*.pdf"))):
        doc = fitz.open(path)
        blanks, where = [], None
        for i, pg in enumerate(doc):
            t = pg.get_text().strip()
            if not t:
                blanks.append(i + 1)
            if "SECTION2" in t and where is None:
                where = i + 1
        print(f"{os.path.basename(path)[:-4]:<22} {len(doc):>5}  {str(where):>8}  {blanks}")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    {"gen": gen, "measure": measure, "read": read}[cmd]()
