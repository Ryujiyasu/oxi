# -*- coding: utf-8 -*-
"""What does the keep-test count -- the line's OWN notes, or only the earlier ones?

Two models explain `_pb_fnarea` (Word keeps R07 and rolls NOTE7) equally well:

  exclude-own    body limit = margin - sep - SUM(notes committed BEFORE this line)
  full-reserve   body limit = margin - sep - SUM(all notes incl. this line's own)

They differ by exactly one note height per own reference, so sweep the final
line's position in fine steps and read the flip point for nown = 1, 2, 3. The
shift of the flip per added own reference IS the coefficient: 0 = exclude-own,
one note height = full-reserve.

An exact-line spacer above the reference block slides everything below it.

    python tools/metrics/_pb_fnkeep_gen.py [--sweep lo hi step]
    python tools/metrics/_pb_fnkeep_read.py word
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fndefer_gen import CT, RELS, DRELS, SETTINGS, STYLES, SEP

# The probe base declares the separator as w:id="0" and the continuation as
# w:id="1"; the OOXML convention is -1 and 0. FNK_SEPID=1 restores the
# convention so the separator reservation can be read against it.
SEPID = os.environ.get("FNK_SEPID", "0")
if SEPID == "1":
    SEP = SEP.replace('w:type="separator" w:id="0"',
                      'w:type="separator" w:id="-1"')
    SEP = SEP.replace('w:type="continuationSeparator" w:id="1"',
                      'w:type="continuationSeparator" w:id="0"')
    # settings.xml names the special footnotes by id, so it has to follow.
    SETTINGS = SETTINGS.replace('<w:footnote w:id="0"/><w:footnote w:id="1"/>',
                                '<w:footnote w:id="-1"/><w:footnote w:id="0"/>')

NPRIOR = int(os.environ.get("FNK_PRIOR", "2"))
FNSZ = int(os.environ.get("FNK_FNSZ", "20"))      # footnote size, half-points
NFILL = int(os.environ.get("FNK_FILL", "43"))
OUT = r"C:\tmp\pb_fnkeep_p%d_z%d_f%d_s%s" % (NPRIOR, FNSZ, NFILL, SEPID)
NOWN = [1, 2, 3]
NOTE = ("Note %d: a single line of footnote text, long enough to fill one line "
        "of the note area but not two. ")


def build(tag, spacer_tw, nown):
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
    for k in range(nown):
        runs += ('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                 '<w:footnoteReference w:id="%d"/></w:r>' % nid)
        ids.append(nid)
        nid += 1
    body += "<w:p>" + runs + "</w:p>"
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
                  'w:lineRule="auto"/><w:rPr><w:sz w:val="%d"/></w:rPr></w:pPr>'
                  '<w:r><w:rPr><w:sz w:val="%d"/></w:rPr>'
                  '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:footnote>'
                  % (fid, FNSZ, FNSZ, NOTE % (n + 1)))
    footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                 + SEP + notes + "</w:footnotes>")
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/footnotes.xml", footnotes)
        z.writestr("word/document.xml", doc)
    return path


def arms(sweep):
    return [("s%05d_o%d" % (x, o), x, o) for x in sweep for o in NOWN]


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
    for t, x, o in a:
        build(t, x, o)
    print("built %d arms (sweep %d..%d) in %s" % (len(a), sw[0], sw[-1], OUT))
