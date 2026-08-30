# -*- coding: utf-8 -*-
"""The `_pb_fnkeep` sweep re-run in `_pb_fnarea`'s geometry.

`_pb_fnkeep` (A4, Calibri 11 body, 10pt notes) put the keep boundary one note
height lower per own reference -- the line's own notes ARE reserved -- and
showed no roll at all. `_pb_fnarea` (Letter, TNR 12 body, 8pt notes,
widowControl 0) rolls. Run the same spacer sweep in that geometry so the two
are read by one instrument and the discriminator is whatever is left.

    python tools/metrics/_pb_fnkeep2_gen.py [--sweep lo hi step]
    python tools/metrics/_pb_fnkeep2_read.py word
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Flags morph this probe toward `_pb_fnkeep` one group at a time, so the
# attribute that decides "reserve the own note" (A4/Calibri) vs "roll it"
# (Letter/TNR) can be bisected. PLUMB drops the footnote-text style, the
# footnoteRef mark, the ref character style, the trailing text and
# widowControl=0; TYPO switches body font and page size.
# `_pb_fnarea` and this probe ship NO settings.xml, so Word reads them in a
# legacy compatibility mode; `_pb_fnkeep` writes compatibilityMode 15.
# Footnote placement is compat-gated (w:footnoteLayoutLikeWW8), so this is
# the last untested difference between the rolling and non-rolling probes.
COMPAT = os.environ.get("FNK2_COMPAT", "0")   # "0" = ship no settings.xml
PLUMB = os.environ.get("FNK2_PLUMB") == "1"
TYPO = os.environ.get("FNK2_TYPO") == "1"
# TYPO is the pair; FNK2_FONT / FNK2_PAGE split it for the last bisection.
FONT_C = os.environ.get("FNK2_FONT", "1" if TYPO else "0") == "1"
PAGE_A4 = os.environ.get("FNK2_PAGE", "1" if TYPO else "0") == "1"
NFILL = int(os.environ.get("FNK2_FILL", "17"))
NPRIOR = int(os.environ.get("FNK2_PRIOR", "6"))
OUT = (r"C:\tmp\pb_fnkeep2_p%d_f%d_%d%d%d"
       % (NPRIOR, NFILL, PLUMB, FONT_C, PAGE_A4)) + ("_c" + COMPAT if COMPAT != "0" else "")
NOWN = [1, 2, 3]

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
      + ('<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
         if COMPAT != '0' else '') +
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
         + ('<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
            if COMPAT != '0' else '') +
         '</Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings ' + W_NS + '>'
            '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
            '<w:compat><w:compatSetting w:name="compatibilityMode"'
            ' w:uri="http://schemas.microsoft.com/office/word" w:val="' + COMPAT + '"/></w:compat>'
            '</w:settings>')
FONT = 'Calibri' if FONT_C else 'Times New Roman'
BODYSZ = 22 if FONT_C else 24
PGSZ = ('<w:pgSz w:w="11906" w:h="16838"/>' if PAGE_A4
        else '<w:pgSz w:w="12240" w:h="15840"/>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles ' + W_NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="%s" w:hAnsi="%s"/><w:sz w:val="%d"/>' % (FONT, FONT, BODYSZ) +
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          + ('<w:pPr/>' if PLUMB else '<w:pPr><w:widowControl w:val="0"/></w:pPr>') +
          '</w:style>'
          + ('' if PLUMB else
             '<w:style w:type="paragraph" w:styleId="FnText"><w:name w:val="footnote text"/>'
             '<w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>'
             '<w:style w:type="character" w:styleId="FnRef"><w:name w:val="footnote reference"/>'
             '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>') +
          '</w:styles>')
SEP = ('<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" '
       'w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
       '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing '
       'w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
       '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>')
FILLER = ('Lorem ipsum dolor sit amet consectetur adipiscing elit sed do '
          'eiusmod tempor incididunt ut labore et dolore.')


def ref_run(fid):
    rpr = ('<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>' if PLUMB
           else '<w:rPr><w:rStyle w:val="FnRef"/></w:rPr>')
    return '<w:r>%s<w:footnoteReference w:id="%d"/></w:r>' % (rpr, fid)


TAIL = "" if PLUMB else '<w:r><w:t xml:space="preserve">; and more</w:t></w:r>'


def note_xml(fid):
    if PLUMB:
        return ('<w:footnote w:id="%d"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
                'w:lineRule="auto"/><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:sz w:val="16"/></w:rPr>'
                '<w:t xml:space="preserve"> NOTE%d ref text.</w:t></w:r></w:p>'
                '</w:footnote>' % (fid, fid))
    return ('<w:footnote w:id="%d"><w:p><w:pPr><w:pStyle w:val="FnText"/></w:pPr>'
            '<w:r><w:rPr><w:rStyle w:val="FnRef"/></w:rPr><w:footnoteRef/></w:r>'
            '<w:r><w:t xml:space="preserve"> NOTE%d ref text.</w:t></w:r></w:p>'
            '</w:footnote>' % (fid, fid))


def build(tag, spacer_tw, nown):
    body = ""
    for k in range(NFILL):
        body += ('<w:p><w:r><w:t xml:space="preserve">F%02d %s</w:t></w:r></w:p>'
                 % (k, FILLER))
    body += ('<w:p><w:pPr><w:spacing w:line="%d" w:lineRule="exact"/></w:pPr>'
             '<w:r><w:t xml:space="preserve">SPACER</w:t></w:r></w:p>' % spacer_tw)
    nid, ids = 1, []
    for i in range(NPRIOR):
        body += ('<w:p><w:r><w:t xml:space="preserve">R%02d short reference line with a '
                 'citation</w:t></w:r>%s%s</w:p>' % (i + 1, ref_run(nid), TAIL))
        ids.append(nid)
        nid += 1
    runs = '<w:r><w:t xml:space="preserve">FINAL line with citations</w:t></w:r>'
    for k in range(nown):
        runs += ref_run(nid) + TAIL
        ids.append(nid)
        nid += 1
    body += "<w:p>" + runs + "</w:p>"
    for i in range(8):
        body += ('<w:p><w:r><w:t xml:space="preserve">TAIL%03d plain line</w:t></w:r></w:p>'
                 % (i + 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document ' + W_NS + '><w:body>' + body +
           '<w:sectPr>' + PGSZ +
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    notes = ""
    for fid in ids:
        notes += note_xml(fid)
    footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<w:footnotes ' + W_NS + '>' + SEP + notes + '</w:footnotes>')
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/footnotes.xml", footnotes)
        if COMPAT != "0":
            z.writestr("word/settings.xml", SETTINGS)
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
    n = 0
    for x in sw:
        for o in NOWN:
            build("s%05d_o%d" % (x, o), x, o)
            n += 1
    print("built %d arms (sweep %d..%d) in %s" % (n, sw[0], sw[-1], OUT))
