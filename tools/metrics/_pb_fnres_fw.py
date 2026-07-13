"""Framework-geometry fn reservation calibration (Calibri 11, Normal after=0,
FootnoteText after=60, footnotePr declares -1,0 — no continuationNotice).

Variants: A (no fn) / B (1 fn on para 2) / C (2 fns: paras 2 and 4).
Same-transition flip differences give R1 (sep-region + fn box) and
R2-R1 (2nd fn box + inter-note gap) at 0.1pt.

Usage:
  python _pb_fnres_fw.py gen A:100:1100:50 B:... C:...
  python _pb_fnres_gen.py measure "fnr_FA_*"   (etc.)
"""
import os, sys, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_fnres_gen as g

CT = g.CT
RELS = g.RELS
DOCRELS = g.DOCRELS

SETTINGS = (g.SETTINGS15
            .replace('w:val="15"', 'w:val="15"')
            .replace('</w:compat>',
                     '</w:compat><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'))

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {g.W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
          '<w:lang w:val="en-GB" w:eastAsia="en-GB" w:bidi="ar-SA"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:spacing w:after="60"/></w:pPr>'
          '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
          '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/>'
          '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
          '</w:styles>')

FN_A = 'Corporate governance code guidance published on the departmental website for reference.'
FN_B = 'Code of conduct for board members of public bodies published for reference purposes.'

FOOTNOTES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<w:footnotes {g.W_NS}>'
             '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
             '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
             '<w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
             '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
             f'<w:r><w:t xml:space="preserve"> {FN_A}</w:t></w:r></w:p></w:footnote>'
             '<w:footnote w:id="3"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
             '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
             f'<w:r><w:t xml:space="preserve"> {FN_B}</w:t></w:r></w:p></w:footnote>'
             '</w:footnotes>')


def para(i, ref_id=None):
    ref = (f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
           f'<w:footnoteReference w:id="{ref_id}"/></w:r>') if ref_id else ''
    return f'<w:p><w:r><w:t>Item {i:02d} alpha beta gamma.</w:t></w:r>{ref}</w:p>'


def build(bottom_tw, variant, n=60):
    paras = []
    for i in range(n):
        rid = None
        if variant in ('B', 'C') and i == 1:
            rid = 2
        if variant == 'C' and i == 3:
            rid = 3
        paras.append(para(i + 1, rid))
    body = ''.join(paras)
    body += (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1440" w:right="851" w:bottom="{bottom_tw}" '
             f'w:left="851" w:header="709" w:footer="709" w:gutter="0"/>'
             f'<w:docGrid w:linePitch="360"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {g.W_NS}><w:body>{body}</w:body></w:document>')


def emit(variant, xs):
    os.makedirs(g.OUTDIR, exist_ok=True)
    for x in xs:
        with zipfile.ZipFile(os.path.join(g.OUTDIR, f'fnr_F{variant}_{x:04d}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/settings.xml', SETTINGS)
            z.writestr('word/footnotes.xml', FOOTNOTES)
            z.writestr('word/document.xml', build(x, variant))
    print('emitted', variant, len(list(xs)))


if __name__ == '__main__':
    for spec in sys.argv[1:]:
        parts = spec.split(':')
        emit(parts[0], range(int(parts[1]), int(parts[2]) + 1, int(parts[3])))
