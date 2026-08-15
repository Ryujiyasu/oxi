# -*- coding: utf-8 -*-
"""What does Word substitute for a font that is not installed, and what decides it?

_pb_linepitch_gen.py showed two different uninstalled faces -- `Humnst777 Lt BT`
and the invented `Grammarsaurus` -- both coming back at 1.17257em, Cambria's
natural height, when the probe declared them inline with NO fontTable part.  But
S811 measured Word substituting Humnst777 with CALIBRI in ukframework, a real
document that does carry a fontTable entry.  Both cannot be one rule, so the
fontTable entry (and the PANOSE it carries) has to be the discriminator.

Rather than infer the substitute from the line height, read it directly: the
exported PDF names the face it actually embedded for each span.

Arms hold the run's font name fixed and vary only what the fontTable says about
it -- absent, present but bare, present with a family class, present with the
real PANOSE, and present with a PANOSE borrowed from a different classification.

  python _pb_fontsub2_gen.py gen
  python _pb_fontsub2_gen.py pdf      # Word truth: embedded face + line pitch
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_fontsub2")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = "Humnst777 Lt BT"
# Humanist 777 Light's own PANOSE (swiss / light / ...), as the BT foundry ships it
PANOSE_SANS = "020B0402020203020204"
PANOSE_ROMAN = "02040503050406030204"      # a Times-like serif classification
ARMS = [
    ("a_no_fonttable", None),
    ("b_bare_entry", '<w:font w:name="%s"/>' % FACE),
    ("c_family_swiss",
     '<w:font w:name="%s"><w:family w:val="swiss"/></w:font>' % FACE),
    ("d_family_roman",
     '<w:font w:name="%s"><w:family w:val="roman"/></w:font>' % FACE),
    ("e_panose_sans",
     '<w:font w:name="%s"><w:panose1 w:val="%s"/><w:family w:val="swiss"/>'
     "</w:font>" % (FACE, PANOSE_SANS)),
    ("f_panose_roman",
     '<w:font w:name="%s"><w:panose1 w:val="%s"/><w:family w:val="roman"/>'
     "</w:font>" % (FACE, PANOSE_ROMAN)),
    # a-f all came back Cambria, so nothing the fontTable says decides it. S811
    # measured CALIBRI for this same face in ukframework, a document that has a
    # theme part -- so test the theme's minor Latin font as the discriminator.
    ("g_theme_calibri", '<w:font w:name="%s"/>' % FACE, "Calibri"),
    ("h_theme_georgia", '<w:font w:name="%s"/>' % FACE, "Georgia"),
    ("i_theme_only_no_ft", None, "Calibri"),
    # a-i all name the missing face DIRECTLY in rFonts. S811's document reaches it
    # through the theme (`w:asciiTheme="minorHAnsi"` over a theme whose minor
    # latin IS the missing face), which is a different lookup: Word resolves the
    # theme slot first and can fall back to the theme's own default.
    ("j_theme_resolved", '<w:font w:name="%s"/>' % FACE, FACE, "theme"),
    # a-j all land on Cambria, yet ukframework's own PDF embeds CALIBRI for this
    # same face. Its fontTable entry is richer than anything above: PANOSE
    # 020B0402030504020204 (different digits from the foundry value guessed for
    # arm e), plus charset / pitch / the <w:sig> Unicode-range signature. Copied
    # verbatim below -- if Word then picks Calibri, entry completeness is what
    # enables the PANOSE match.
    ("k_ukframework_entry", '<w:font w:name="Humnst777 Lt BT"><w:panose1 w:val="020B0402030504020204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="800000AF" w:usb1="1000204A" w:usb2="00000000" w:usb3="00000000" w:csb0="00000011" w:csb1="00000000"/></w:font>'),
    # l: the real document also DECLARES Calibri in the same fontTable --
    #    test whether Word prefers a face the document already names.
    ("l_plus_calibri_entry", '<w:font w:name="Humnst777 Lt BT"><w:panose1 w:val="020B0402030504020204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="800000AF" w:usb1="1000204A" w:usb2="00000000" w:usb3="00000000" w:csb0="00000011" w:csb1="00000000"/></w:font><w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="E4002EFF" w:usb1="C000247B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font>'),
    # m: the probe has no settings.xml at all, so Word opens it outside any
    #    declared compatibility mode; ukframework declares 15.
    ("m_compat15", '<w:font w:name="Humnst777 Lt BT"><w:panose1 w:val="020B0402030504020204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="800000AF" w:usb1="1000204A" w:usb2="00000000" w:usb3="00000000" w:csb0="00000011" w:csb1="00000000"/></w:font>', None, "compat"),
    # n/o: the probe declares no docDefaults, so Word has no document
    #      default font to fall back to. Vary that default and see whether
    #      the substitute follows it.
    ("n_default_calibri", '<w:font w:name="Humnst777 Lt BT"><w:panose1 w:val="020B0402030504020204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="800000AF" w:usb1="1000204A" w:usb2="00000000" w:usb3="00000000" w:csb0="00000011" w:csb1="00000000"/></w:font>', None, "dd:Calibri"),
    ("o_default_georgia", '<w:font w:name="Humnst777 Lt BT"><w:panose1 w:val="020B0402030504020204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="800000AF" w:usb1="1000204A" w:usb2="00000000" w:usb3="00000000" w:csb0="00000011" w:csb1="00000000"/></w:font>', None, "dd:Georgia"),
]
THEME = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' name="probe"><a:themeElements><a:clrScheme name="p">'
    + "".join('<a:%s><a:srgbClr val="000000"/></a:%s>' % (t, t) for t in
              ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
               "accent4", "accent5", "accent6", "hlink", "folHlink"))
    + '</a:clrScheme><a:fontScheme name="p">'
    '<a:majorFont><a:latin typeface="%s"/><a:ea typeface=""/><a:cs typeface=""/>'
    "</a:majorFont>"
    '<a:minorFont><a:latin typeface="%s"/><a:ea typeface=""/><a:cs typeface=""/>'
    "</a:minorFont></a:fontScheme>"
    '<a:fmtScheme name="p"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/>'
    "</a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
    '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
    "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>"
    "<a:effectStyle><a:effectLst/></a:effectStyle>"
    "<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>"
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
    "</a:fmtScheme></a:themeElements></a:theme>")
SENT = ("The registrar must determine the percentage of care that a person has "
        "for a child during a care period and notify each person concerned. ")


def docx(ai):
    return os.path.join(OUT, "fontsub2_%d.docx" % ai)


def gen():
    os.makedirs(OUT, exist_ok=True)
    for ai, arm in enumerate(ARMS):
        name, fentry = arm[0], arm[1]
        theme = arm[2] if len(arm) > 2 else None
        via_theme = len(arm) > 3 and arm[3] == "theme"
        compat = len(arm) > 3 and arm[3] == "compat"
        dd_font = arm[3][3:] if len(arm) > 3 and str(arm[3]).startswith("dd:") else None
        # one arm per FILE: a fontTable is per-document, so the arms cannot share
        # one document the way page-per-arm probes usually do
        body = (
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr><w:r><w:rPr>'
            "%s<w:sz w:val=\"20\"/></w:rPr>"
            "<w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>"
            % ('<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>'
               if via_theme else
               '<w:rFonts w:ascii="%s" w:hAnsi="%s"/>' % (FACE, FACE), SENT * 8))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
               "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11907" w:h="16839"/>'
               '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
               'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        dd_xml = ("<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"%s\""
                  " w:hAnsi=\"%s\" w:eastAsia=\"%s\" w:cs=\"%s\"/></w:rPr>"
                  "</w:rPrDefault></w:docDefaults>"
                  % (dd_font, dd_font, dd_font, dd_font)) if dd_font else ""
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles '
                  + NS + ">" + dd_xml
                  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
                  '<w:name w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr>'
                  "</w:style></w:styles>")
        with zipfile.ZipFile(docx(ai), "w", zipfile.ZIP_DEFLATED) as z:
            ct, drels = CT, DRELS
            if fentry:
                ct = CT.replace(
                    "</Types>",
                    '<Override PartName="/word/fontTable.xml" ContentType='
                    '"application/vnd.openxmlformats-officedocument.'
                    'wordprocessingml.fontTable+xml"/></Types>')
                drels = DRELS.replace(
                    "</Relationships>",
                    '<Relationship Id="rIdFT" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/fontTable" '
                    'Target="fontTable.xml"/></Relationships>')
                z.writestr("word/fontTable.xml",
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                           "<w:fonts " + NS + ">" + fentry + "</w:fonts>")
            # declare the theme BEFORE writing the manifest parts -- an
            # undeclared part is invisible to Word and the arm silently
            # degenerates into the no-theme case
            if theme:
                ct = ct.replace(
                    "</Types>",
                    '<Override PartName="/word/theme/theme1.xml" ContentType='
                    '"application/vnd.openxmlformats-officedocument.theme+xml"/>'
                    "</Types>")
                drels = drels.replace(
                    "</Relationships>",
                    '<Relationship Id="rIdTh" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/theme" '
                    'Target="theme/theme1.xml"/></Relationships>')
                z.writestr("word/theme/theme1.xml", THEME % (theme, theme))
            if compat:
                ct = ct.replace(
                    "</Types>",
                    '<Override PartName="/word/settings.xml" ContentType='
                    '"application/vnd.openxmlformats-officedocument.'
                    'wordprocessingml.settings+xml"/></Types>')
                drels = drels.replace(
                    "</Relationships>",
                    '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/settings" '
                    'Target="settings.xml"/></Relationships>')
                z.writestr("word/settings.xml",
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                           "<w:settings " + NS + "><w:compat>"
                           '<w:compatSetting w:name="compatibilityMode"'
                           ' w:uri="http://schemas.microsoft.com/office/word"'
                           ' w:val="15"/></w:compat></w:settings>')
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/document.xml", doc)
    print("wrote", len(ARMS), "arms to", OUT)


def pdf():
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    print("%-16s %-26s %9s %9s" % ("arm", "embedded face", "pitch", "em"))
    try:
        for ai, arm in enumerate(ARMS):
            name = arm[0]
            out = docx(ai).replace(".docx", ".pdf")
            d = app.Documents.Open(docx(ai), ReadOnly=True)
            try:
                d.ExportAsFixedFormat(out, 17)
            finally:
                d.Close(False)
            doc = fitz.open(out)
            face, ys = "?", []
            for bl in doc[0].get_text("dict")["blocks"]:
                for ln in bl.get("lines", []):
                    for sp in ln["spans"]:
                        if sp["text"].strip():
                            face = sp["font"]
                            ys.append(round(sp["origin"][1], 3))
                            break
            ys = sorted(set(ys))
            pitch = (ys[-1] - ys[0]) / (len(ys) - 1) if len(ys) > 1 else 0.0
            print("%-16s %-26s %9.4f %9.5f" % (name, face, pitch, pitch / 10.0))
    finally:
        app.Quit()


if __name__ == "__main__":
    if sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
