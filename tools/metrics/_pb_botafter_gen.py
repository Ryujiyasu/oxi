# -*- coding: utf-8 -*-
"""Does a page bottom demand room for the paragraph's trailing space?

On `reports__0018715b4769984f` Word's page-5 limit measures 8.0pt stricter than
Oxi's, and 8.0pt is that document's docDefaults `w:after`. If Word really counts
a paragraph's space-after when deciding whether its last line fits, the rule is
not specific to a footnote boundary and should show on a plain page too.

One paragraph at the page bottom, its position swept by an exact-line spacer,
its `w:after` swept as the second axis. If the after-space is in the test the
flip position moves DOWN by exactly the after value; if it is not, the flip
stays put for every after.

    python tools/metrics/_pb_botafter_gen.py [--sweep lo hi step]
    python tools/metrics/_pb_botafter_read.py word|oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fndefer_gen import CT, RELS, DRELS, SETTINGS, STYLES, SEP

# This probe carries no notes, so the footnotes part must go from both the
# content types and the document rels or Word rejects the package.
import re as _re
if os.environ.get("BA_LEGACY") == "1":
    CT = _re.sub(r'<Override PartName="/word/settings[^>]*>', "", CT)
    DRELS = _re.sub(r'<Relationship [^>]*settings\.xml"\s*/>', "", DRELS)
if os.environ.get("BA_NOTES") != "1":
    CT = _re.sub(r'<Override PartName="/word/footnotes[^>]*>', "", CT)
    DRELS = _re.sub(r'<Relationship [^>]*footnotes\.xml"\s*/>', "", DRELS)
    SETTINGS = _re.sub(r"<w:footnotePr>.*?</w:footnotePr>", "", SETTINGS, flags=_re.S)
NOTE = ("Note %d: a single line of footnote text, long enough to fill one line "
        "of the note area but not two. ")

# BA_NOTES=1 puts two one-note reference lines above the tested paragraph, so
# the page foot is spoken for. Word counts the trailing space at a page bottom
# only when it is: the plain sweep flips at the same spacer for every after.
# BA_LEGACY=1 ships no settings.xml, i.e. the old compatibility mode. S1244
# showed footnote placement is compat-gated, so the trailing-space term has
# to be read in both regimes before it can ship.
LEGACY = os.environ.get("BA_LEGACY") == "1"
NOTES = os.environ.get("BA_NOTES") == "1"
OUT = (r"C:\tmp\pb_botafter" + ("_fn" if NOTES else "")
       + ("_legacy" if LEGACY else ""))
NFILL = int(os.environ.get("BA_FILL", "46"))
AFTERS = [0, 160, 320]          # twips: 0, 8pt, 16pt
MULT = os.environ.get("BA_MULT", "240")   # w:line on the tested paragraph


def build(tag, spacer_tw, after_tw):
    body = ""
    for i in range(NFILL):
        body += ('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                 '<w:r><w:t xml:space="preserve">FILL%03d line</w:t></w:r></w:p>'
                 % (i + 1))
    body += ('<w:p><w:pPr><w:spacing w:after="0" w:line="%d" w:lineRule="exact"/></w:pPr>'
             '<w:r><w:t xml:space="preserve">SPACER</w:t></w:r></w:p>' % spacer_tw)
    if NOTES:
        for k in range(2):
            body += ('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                     '<w:r><w:t xml:space="preserve">PRIOR%02d line</w:t></w:r>'
                     '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                     '<w:footnoteReference w:id="%d"/></w:r></w:p>' % (k + 1, k + 2))
    body += ('<w:p><w:pPr><w:spacing w:after="%d" w:line="%s" w:lineRule="auto"/></w:pPr>'
             '<w:r><w:t xml:space="preserve">FINAL line</w:t></w:r></w:p>'
             % (after_tw, MULT))
    for i in range(8):
        body += ('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                 '<w:r><w:t xml:space="preserve">TAIL%03d line</w:t></w:r></w:p>'
                 % (i + 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           "<w:body>\n" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
           ' w:header="708" w:footer="708" w:gutter="0"/>'
           "</w:sectPr></w:body></w:document>")
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        if not LEGACY:
            z.writestr("word/settings.xml", SETTINGS)
        if NOTES:
            notes = "".join(
                '<w:footnote w:id="%d"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
                'w:lineRule="auto"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
                '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:footnote>'
                % (k + 2, NOTE % (k + 1)) for k in range(2))
            z.writestr("word/footnotes.xml",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/'
                       'wordprocessingml/2006/main">' + SEP + notes + "</w:footnotes>")
        z.writestr("word/document.xml", doc)
    return path


def parse_sweep(argv):
    if "--sweep" in argv:
        i = argv.index("--sweep")
        return list(range(int(argv[i + 1]), int(argv[i + 2]) + 1, int(argv[i + 3])))
    return list(range(200, 641, 40))


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    sw = parse_sweep(sys.argv)
    n = 0
    for x in sw:
        for a in AFTERS:
            build("s%05d_a%04d" % (x, a), x, a)
            n += 1
    print("built %d arms (sweep %d..%d) in %s" % (n, sw[0], sw[-1], OUT))
