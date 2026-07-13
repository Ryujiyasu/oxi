"""J/K bullet-body variants for the fn-reservation probe (_pb_fnres_gen).

J = bullets + footnote ref, K = bullets only. The bullet marker is the real
Symbol  (written via a python escape — the bash-heredoc F0B7-loss trap
struck the inline version and produced two INCONSISTENT batches).

Usage:
  python _pb_fnres_jk.py gen J:100:1100:50  (etc.)
  python _pb_fnres_gen.py measure "fnr_J_*"
"""
import os, sys, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_fnres_gen as g

FOOTER = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:ftr {g.W_NS}>'
          '<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
          '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
          '<w:r><w:t>4</w:t></w:r>'
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
          '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>'
          '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>'
          '</w:ftr>')

FOOTER_STYLE = ('<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>'
                '<w:pPr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs>'
                '<w:spacing w:line="260" w:lineRule="exact"/></w:pPr></w:style>')

STYLES_D = g.STYLES.replace('</w:styles>', FOOTER_STYLE + '</w:styles>')

CT_J = g.CT.replace('</Types>',
  '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
  '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')

DOCRELS_J = g.DOCRELS.replace('</Relationships>',
  '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
  '<Relationship Id="rId21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')

BULLET = ''
NUMBERING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<w:numbering {g.W_NS}>'
             '<w:abstractNum w:abstractNumId="12"><w:multiLevelType w:val="hybridMultilevel"/>'
             '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
             f'<w:lvlText w:val="{BULLET}"/><w:lvlJc w:val="left"/>'
             '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
             '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'
             '<w:num w:numId="21"><w:abstractNumId w:val="12"/></w:num>'
             '</w:numbering>')


def bpara(i, with_ref=False):
    ref = ('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
           '<w:footnoteReference w:id="2"/></w:r>') if with_ref else ''
    return ('<w:p><w:pPr><w:widowControl/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="21"/></w:numPr>'
            '<w:spacing w:before="0" w:after="0"/></w:pPr>'
            f'<w:r><w:t>Item {i:02d} alpha beta gamma.</w:t></w:r>{ref}</w:p>')


def build(spacer_tw, with_fn):
    paras = []
    for i in range(52):
        paras.append(bpara(i + 1, with_ref=(with_fn and i == 1)))
        if i == 2:
            paras.append(f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{spacer_tw}" w:lineRule="exact"/></w:pPr></w:p>')
    body = ''.join(paras)
    body += (f'<w:sectPr><w:footerReference w:type="default" r:id="rId20"/><w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1440" w:right="851" w:bottom="1620" '
             f'w:left="851" w:header="709" w:footer="709" w:gutter="0"/>'
             f'<w:docGrid w:linePitch="360"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {g.W_NS}><w:body>{body}</w:body></w:document>')


def emit(tag, with_fn, xs):
    os.makedirs(g.OUTDIR, exist_ok=True)
    for x in xs:
        with zipfile.ZipFile(os.path.join(g.OUTDIR, f'fnr_{tag}_{x:04d}.docx'), 'w',
                             zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT_J)
            z.writestr('_rels/.rels', g.RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS_J)
            z.writestr('word/styles.xml', STYLES_D)
            z.writestr('word/settings.xml', g.SETTINGS15)
            z.writestr('word/footnotes.xml', g.FOOTNOTES)
            z.writestr('word/footer1.xml', FOOTER)
            z.writestr('word/numbering.xml', NUMBERING)
            z.writestr('word/document.xml', build(x, with_fn))
    print('emitted', tag, len(list(xs)))


if __name__ == '__main__':
    for spec in sys.argv[1:]:
        tag, lo, hi, step = spec.split(':')
        emit(tag, tag == 'J', range(int(lo), int(hi) + 1, int(step)))
