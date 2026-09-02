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

# 2026-09-02 — does the rule compare LOGICAL parity to PHYSICAL parity, or does
# it preserve the alternation the document's FIRST page established?
#
# Every b-case above restarts sect0 at 1 on physical 1, where odd<->odd makes
# the two readings identical -- so the b-round could not tell them apart.
# reference__0b6f3b32 is the discriminator: sect0 restarts at 272 (EVEN) on
# physical 1, sect1 restarts at 272 again, and Word blanks page 2. Under
# "logical parity == physical parity" (what S957 implements) 272 onto physical
# 2 MATCHES and no pad is due; under "keep alternating" it does not.
#
#   c1  S957 says no blank   | alternation says blank      <- decides it
#   c2  S957 says blank      | alternation says no blank   <- decides it, other way
#   c3  control: no evenAndOddHeaders, so neither pads
#   c4  control: odd cover, the classic case, both agree -> blank
CASES += [
    ("c1_evencover_restart_even", True, 'w:start="272"', 'w:start="272"', True, True),
    ("c2_evencover_restart_odd",  True, 'w:start="272"', 'w:start="273"', True, True),
    ("c3_noeoh_evencover",        False, 'w:start="272"', 'w:start="272"', True, True),
    ("c4_oddcover_restart_odd",   True, 'w:start="271"', 'w:start="271"', True, True),
]

# Three-section arms: what does `oddPage` count -- physical pages, or LOGICAL
# ones? In 0b6f3b32 the oddPage section restarts at 273 (odd) and Word puts it
# on physical page 4 (EVEN), which a physical reading cannot produce. Fixing the
# blank-page rule without settling this one would make S732 insert a WRONG blank
# right after the newly-correct one.
#   (id, eoh, sect0_pgnum, sect1_pgnum, sect2_type, sect2_pgnum)

# 2026-09-03 -- a CONTINUOUS section that declares its own pgNumType start.
# reference__0ea3ec86 has two (sec2 start=90, sec3 start=88) and Word blanks
# its page 2, so the parity rule fires for a continuous section too. What is
# NOT known from that document alone is whether such a section always begins a
# PAGE (a page can only carry one number, so a restart cannot take effect
# mid-page) or only does so when the parity padding forces it.
#   g1  parity CONFLICTS -> expect a blank page and sec2 at the top of page 3
#   g2  parity AGREES    -> the discriminator: does sec2 still open page 2, or
#                           does its content continue mid-page 1?
CASES_CONT = [
    ("g1_cont_restart_conflict", True, 'w:start="1"', 'w:start="1"'),
    ("g2_cont_restart_agree",    True, 'w:start="1"', 'w:start="2"'),
    ("g3_cont_restart_noeoh",    False, 'w:start="1"', 'w:start="1"'),
]

# 2026-09-03 -- the g arms above padded NOTHING, yet reference__0ea3ec86 (whose
# sec2 is also continuous with a restart) gets a blank page 2. The difference is
# that 0ea3ec86's cover FILLS page 1, so its continuous section has to BEGIN a
# page; in the g arms the cover is one line and the section just continues on
# it, so no page ever starts and no parity question is asked.
#   h arms: the cover is padded to fill page 1, so sec2 must open page 2.
#   h1  conflict -> expect a blank page 2, sec2 on page 3
#   h2  agree    -> expect no blank, sec2 on page 2
# Sweep the cover length instead of guessing it: the arm that matters is the
# first one where the cover fills page 1, and 46 lines of 12pt Arial did not.
CASES_FILL = ([(f"h1c{n}_conflict", True, 'w:start="1"', 'w:start="3"', n)
               for n in range(46, 58, 2)]
              + [(f"h2c{n}_agree", True, 'w:start="1"', 'w:start="2"', n)
                 for n in range(46, 58, 2)])

CASES3 = [
    ("e1_evencover_oddpage_r273", True, 'w:start="272"', 'w:start="272"',
     "oddPage", 'w:start="273"'),
    ("e2_evencover_oddpage_norestart", True, 'w:start="272"', 'w:start="272"',
     "oddPage", None),
    ("e3_oddcover_oddpage_norestart", True, 'w:start="1"', None,
     "oddPage", None),
    ("e4_evencover_evenpage_norestart", True, 'w:start="272"', 'w:start="272"',
     "evenPage", None),
]

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


def sect(pgnum, titlepg, inner=True, stype=None):
    """A sectPr; `inner` wraps it in the trailing paragraph of the section."""
    body = ""
    if stype:
        body += f'<w:type w:val="{stype}"/>'   # must precede pgSz per the schema
    body += PGSZ
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
        write(cid, doc, eoh)
        print("gen", cid)
    for (cid, eoh, pg0, pg1, nfill) in CASES_FILL:
        cover = "".join(para(f"COVERLINE{i:02d}") for i in range(nfill))
        body = (cover + sect(pg0, True)
                + para("SECTION2 FIRST LINE") + para("more text")
                + sect(pg1, True, inner=False, stype="continuous"))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
               f'<w:body>{body}</w:body></w:document>')
        write(cid, doc, eoh)
        print("gen", cid)
    for (cid, eoh, pg0, pg1) in CASES_CONT:
        # sec2 is CONTINUOUS and restarts numbering. Its marker paragraph is
        # SECTION2 so `read` reports which page it landed on.
        body = (para("COVER PAGE") + sect(pg0, True)
                + para("SECTION2 FIRST LINE") + para("more text")
                + sect(pg1, True, inner=False, stype="continuous"))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
               f'<w:body>{body}</w:body></w:document>')
        write(cid, doc, eoh)
        print("gen", cid)
    for (cid, eoh, pg0, pg1, s2type, pg2) in CASES3:
        body = (para("COVER PAGE") + sect(pg0, True)
                + para("SECTION2 FIRST LINE") + sect(pg1, True)
                + para("SECTION3 FIRST LINE")
                + sect(pg2, True, inner=False, stype=s2type))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
               f'<w:body>{body}</w:body></w:document>')
        write(cid, doc, eoh)
        print("gen", cid)


def write(cid, doc, eoh):
    with zipfile.ZipFile(os.path.join(OUTDIR, cid + ".docx"), "w",
                         zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOCRELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", settings_xml(eoh))


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
    print(f"{'case':<32} {'pages':>5}  {'sect2':>5} {'sect3':>5}  blanks")
    for path in sorted(glob.glob(os.path.join(OUTDIR, "*.pdf"))):
        doc = fitz.open(path)
        blanks, where, where3 = [], None, None
        for i, pg in enumerate(doc):
            t = pg.get_text().strip()
            if not t:
                blanks.append(i + 1)
            if "SECTION2" in t and where is None:
                where = i + 1
            if "SECTION3" in t and where3 is None:
                where3 = i + 1
        print(f"{os.path.basename(path)[:-4]:<32} {len(doc):>5}  {str(where):>5} "
              f"{str(where3):>5}  {blanks}")


def oxi():
    """The same census off Oxi's own layout, so the two read identically."""
    import json, subprocess, tempfile
    exe = os.environ.get("OXI_GDI_EXE") or os.path.join(
        REPO, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
    print(f"{'case':<32} {'pages':>5}  {'sect2':>5} {'sect3':>5}  blanks")
    for path in sorted(glob.glob(os.path.join(OUTDIR, "*.docx"))):
        with tempfile.TemporaryDirectory() as t:
            dj = os.path.join(t, "l.json")
            r = subprocess.run([exe, os.path.abspath(path), os.path.join(t, "p"),
                                "--dump-layout=" + dj], capture_output=True, timeout=180)
            if r.returncode != 0 or not os.path.exists(dj):
                print(f"{os.path.basename(path)[:-5]:<32}  RENDER FAIL")
                continue
            with open(dj, encoding="utf-8") as f:
                dump = json.load(f)
        blanks, where, where3 = [], None, None
        for i, pg in enumerate(dump["pages"], 1):
            t2 = "".join(e.get("text") or "" for e in pg["elements"]
                         if e.get("type") == "text").strip()
            if not t2:
                blanks.append(i)
            if "SECTION2" in t2 and where is None:
                where = i
            if "SECTION3" in t2 and where3 is None:
                where3 = i
        print(f"{os.path.basename(path)[:-5]:<32} {len(dump['pages']):>5}  "
              f"{str(where):>5} {str(where3):>5}  {blanks}")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    {"gen": gen, "measure": measure, "read": read, "oxi": oxi}[cmd]()
