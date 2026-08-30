# -*- coding: utf-8 -*-
"""At a page bottom with a footnote area, which box does the last line use?

`reports__0018715b4769984f` p5 turns on ~1.8pt: Oxi fits the 3-line paragraph
pi=58 (its last line's natural box ends 1.78pt above the limit) and Word pushes
the whole paragraph. That paragraph has NO footnote reference of its own, yet
sits above a committed note area, and its lines carry a line-spacing multiplier
(effective 14.491 against a natural 13.500).

Two binary questions decide it, and one sweep answers both:

  BOX     does Word measure the last line by its NATURAL box or by its
          MULTIPLIED box? The flip position moves by (mult-1)*natural between
          the two models -- ~3.4pt at mult 1.25, far above the 0.2pt step.
  RELIEF  does the fs/16 fn-boundary relief apply to a paragraph with no own
          refs? (S835's own derivation says that case is "not yet covered".)

Arms: MULT x own-ref-or-not, swept by an exact-line spacer. FINAL is ONE line,
so the flip is the line's own keep test with no widow rule in the way.

    python tools/metrics/_pb_fnbox_gen.py [--sweep lo hi step]
    python tools/metrics/_pb_fnbox_read.py word|oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fndefer_gen import CT, RELS, DRELS, SETTINGS, STYLES, SEP

# S1244 showed footnote placement is compat-gated, so the box rule has to be
# read in both regimes. FNB_LEGACY=1 ships no settings.xml at all (= the old
# compatibility mode); the default writes it, as real modern documents do.
LEGACY = os.environ.get("FNB_LEGACY") == "1"
if LEGACY:
    import re as _re
    CT = _re.sub(r'<Override PartName="/word/settings[^>]*>', "", CT)
    DRELS = _re.sub(r'<Relationship [^>]*settings\.xml"\s*/>', "", DRELS)

OUT = r"C:\tmp\pb_fnbox" + ("_legacy" if LEGACY else "")
NFILL = int(os.environ.get("FNB_FILL", "43"))
NPRIOR = 2
# w:line values for lineRule="auto" (240 = single). 258 ~= the real doc's
# 14.491/13.500; 300 = 1.25 for a signal far above the sweep step.
MULTS = [240, 258, 300]
OWN = [0, 1]          # does FINAL carry a footnote reference of its own?
NOTE = ("Note %d: a single line of footnote text, long enough to fill one line "
        "of the note area but not two. ")

def arms(sweep):
    return [("s%05d_m%d_o%d" % (x, m, o), x, m, o)
            for x in sweep for m in MULTS for o in OWN]


def build(tag, spacer_tw, mult, own):
    body = ""
    for i in range(NFILL):
        body += ('<w:p><w:r><w:t xml:space="preserve">FILL%03d line</w:t></w:r></w:p>'
                 % (i + 1))
    body += ('<w:p><w:pPr><w:spacing w:line="%d" w:lineRule="exact"/></w:pPr>'
             '<w:r><w:t xml:space="preserve">SPACER</w:t></w:r></w:p>' % spacer_tw)
    nid, ids = 2, []
    for k in range(NPRIOR):
        body += ('<w:p><w:r><w:t xml:space="preserve">PRIOR%02d line</w:t></w:r>'
                 '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                 '<w:footnoteReference w:id="%d"/></w:r></w:p>' % (k + 1, nid))
        ids.append(nid)
        nid += 1
    runs = '<w:r><w:t xml:space="preserve">FINAL line</w:t></w:r>'
    if own:
        runs += ('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                 '<w:footnoteReference w:id="%d"/></w:r>' % nid)
        ids.append(nid)
        nid += 1
    body += ('<w:p><w:pPr><w:spacing w:line="%d" w:lineRule="auto"/></w:pPr>%s</w:p>'
             % (mult, runs))
    for i in range(8):
        body += ('<w:p><w:r><w:t xml:space="preserve">TAIL%03d line</w:t></w:r></w:p>'
                 % (i + 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           "<w:body>\n" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
           ' w:header="708" w:footer="708" w:gutter="0"/>'
           "</w:sectPr></w:body></w:document>")
    notes = ""
    for n, fid in enumerate(ids):
        notes += ('<w:footnote w:id="%d"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
                  'w:lineRule="auto"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>'
                  '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
                  '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:footnote>'
                  % (fid, NOTE % (n + 1)))
    footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                 + SEP + notes + "</w:footnotes>")
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        if not LEGACY:
            z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/footnotes.xml", footnotes)
        z.writestr("word/document.xml", doc)
    return path


def parse_sweep(argv):
    if "--sweep" in argv:
        i = argv.index("--sweep")
        lo, hi, st = int(argv[i + 1]), int(argv[i + 2]), int(argv[i + 3])
        return list(range(lo, hi + 1, st))
    return list(range(200, 641, 40))


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    sw = parse_sweep(sys.argv)
    a = arms(sw)
    for t, x, m, o in a:
        build(t, x, m, o)
    print("built %d arms (sweep %d..%d) in %s" % (len(a), sw[0], sw[-1], OUT))
