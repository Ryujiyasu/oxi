# -*- coding: utf-8 -*-
"""Adversarial structural probe generator — hunt for Phase-1 breakers.

Emits a batch of minimal, VALID, multi-page .docx that each stress a layout
path the tuned 87-doc corpus barely exercises:
  vertical writing, multi-column, non-A4 page sizes, landscape, mixed-section
  geometry changes, unusual docGrid, font variants, footnotes, OMML, nested
  tables spanning a page break.

Each doc is written to tools/golden-test/documents/docx/<id>_<name>.docx so the
pagination pipeline picks it up with doc_id = <id>.

Run: python tools/metrics/_probe_gen.py
"""
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx")

MINCHO = "ＭＳ 明朝"
GOTHIC = "ＭＳ ゴシック"

def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# realistic 約物-rich regulation body sentence (wraps to ~2-3 lines each)
SENT = (
    "本項に定める事項については、関係法令及び本規程の趣旨に照らし、"
    "善良なる管理者の注意をもって、誠実かつ適切に取り扱うものとし、"
    "これに関連して生じた一切の事項は、別に定めるところにより、"
    "遅滞なく、かつ、公正に処理しなければならない。"
)

def para(n, font=MINCHO, sz="21", jc="both", extra_ppr="", extra_rpr=""):
    """A numbered regulation paragraph."""
    rpr = (f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
           f'<w:sz w:val="{sz}"/>{extra_rpr}')
    return (
        f'<w:p><w:pPr><w:jc w:val="{jc}"/>{extra_ppr}'
        f'<w:rPr>{rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{rpr}</w:rPr>'
        f'<w:t xml:space="preserve">第{n}条　{esc(SENT)}</w:t></w:r></w:p>'
    )

def body_paras(count, **kw):
    return "".join(para(i + 1, **kw) for i in range(count))

# ---- OOXML skeleton parts -------------------------------------------------

def styles_xml(font=MINCHO, sz="21"):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{sz}"/></w:rPr></w:rPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
        '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>'
        '</w:styles>'
    )

def settings_xml(compat="15", cpunct=True):
    csc = '<w:characterSpacingControl w:val="compressPunctuation"/>' if cpunct else ''
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'{csc}<w:compat><w:compatSetting w:name="compatibilityMode" '
        'w:uri="http://schemas.microsoft.com/office/word" '
        f'w:val="{compat}"/></w:compat></w:settings>'
    )

CT_BASE = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
]

def content_types(extra=()):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            + "".join(CT_BASE) + "".join(extra) + '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

def docrels(extra=()):
    base = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>',
    ]
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(base) + "".join(extra) + '</Relationships>')

DOC_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
          'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

def sectpr(pgsz='<w:pgSz w:w="11906" w:h="16838"/>',
           mar='<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>',
           cols='', textdir='', grid='<w:docGrid w:type="lines" w:linePitch="360"/>', sect_type=''):
    return f'<w:sectPr>{sect_type}{pgsz}{mar}{cols}{textdir}{grid}</w:sectPr>'

def write_docx(path, document, extra_parts=None, ct_extra=(), rel_extra=(),
               font=MINCHO, sz="21", compat="15", cpunct=True):
    parts = {
        "[Content_Types].xml": content_types(ct_extra),
        "_rels/.rels": RELS,
        "word/document.xml": document,
        "word/_rels/document.xml.rels": docrels(rel_extra),
        "word/styles.xml": styles_xml(font, sz),
        "word/settings.xml": settings_xml(compat, cpunct),
    }
    if extra_parts:
        parts.update(extra_parts)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)

def doc(body):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {DOC_NS}><w:body>{body}</w:body></w:document>')

def out(name):
    return os.path.join(DOCX, name)

# ---- the probes -----------------------------------------------------------

def p_vertical():
    """縦書き multi-page. Vertical pagination is basics-only in Oxi."""
    body = body_paras(60) + sectpr(textdir='<w:textDirection w:val="tbRl"/>')
    write_docx(out("probevert_vertical3pg.docx"), doc(body))

def p_twocol():
    """2-column body over several pages. Column flow + pagination."""
    body = body_paras(70) + sectpr(cols='<w:cols w:num="2" w:space="425" w:sep="1"/>')
    write_docx(out("probe2col_twocol.docx"), doc(body))

def p_threecol():
    body = body_paras(80) + sectpr(cols='<w:cols w:num="3" w:space="360" w:sep="1"/>')
    write_docx(out("probe3col_threecol.docx"), doc(body))

def p_b5():
    """B5 page size — different content area → different pagination."""
    body = body_paras(55) + sectpr(pgsz='<w:pgSz w:w="10318" w:h="14570"/>')
    write_docx(out("probeb5_b5body.docx"), doc(body))

def p_landscape():
    body = body_paras(55) + sectpr(
        pgsz='<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>')
    write_docx(out("probeland_landscape.docx"), doc(body))

def p_mixedgeom():
    """Portrait A4 -> landscape A4 -> B5 portrait, section breaks mid-doc."""
    s1 = body_paras(30)
    # section-1 sectPr goes in the last paragraph's pPr
    s1_sect = (f'<w:p><w:pPr>{sectpr()[:-10]}</w:pPr></w:p>')  # not used; build explicitly below
    # Build explicitly: paras then a sectPr-bearing empty para per section boundary
    def sect_break_para(sp):
        return f'<w:p><w:pPr>{sp}</w:pPr></w:p>'
    sec1 = body_paras(28) + sect_break_para(
        sectpr())
    sec2 = body_paras(28) + sect_break_para(
        sectpr(pgsz='<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'))
    sec3 = body_paras(28) + sectpr(pgsz='<w:pgSz w:w="10318" w:h="14570"/>')
    write_docx(out("probemixgeom_sections.docx"), doc(sec1 + sec2 + sec3))

def p_grid_lac():
    """linesAndChars grid, unusual pitch 312 + negative charSpace (atimes-like)."""
    grid = '<w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="-553"/>'
    body = body_paras(60) + sectpr(grid=grid)
    write_docx(out("probelac_linesandchars312.docx"), doc(body))

def p_grid_pitch480():
    """Wide line pitch 480 (double spacing via grid)."""
    grid = '<w:docGrid w:type="lines" w:linePitch="480"/>'
    body = body_paras(40) + sectpr(grid=grid)
    write_docx(out("probep480_pitch480.docx"), doc(body))

def p_yugothic():
    body = body_paras(55, font="游ゴシック") + sectpr()
    write_docx(out("probeyug_yugothic.docx"), doc(body), font="游ゴシック")

def p_meiryo():
    body = body_paras(55, font="メイリオ") + sectpr()
    write_docx(out("probemei_meiryo.docx"), doc(body), font="メイリオ")

def p_linespacing_exact():
    """Exact line spacing 300 (15pt) — lineRule=exact multi-page."""
    ppr = '<w:spacing w:line="300" w:lineRule="exact"/>'
    body = body_paras(50, extra_ppr=ppr) + sectpr()
    write_docx(out("probeexact_exactspacing.docx"), doc(body))

def p_linespacing_mult():
    """Multiple 1.5x line spacing multi-page."""
    ppr = '<w:spacing w:line="360" w:lineRule="auto"/>'
    body = body_paras(45, extra_ppr=ppr) + sectpr(grid='')
    write_docx(out("probemult_mult15.docx"), doc(body))

def p_nogrid_body():
    """No docGrid, MS Mincho body multi-page (LM0 path)."""
    body = body_paras(55) + sectpr(grid='')
    write_docx(out("probenogrid_lm0body.docx"), doc(body))

def p_footnotes():
    """Footnote-heavy multi-page: a footnote on every paragraph."""
    n = 40
    # footnote references in body
    def fn_para(i):
        rpr = f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/><w:sz w:val="21"/>'
        return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{rpr}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">第{i+1}条　{esc(SENT)}</w:t></w:r>'
                f'<w:r><w:rPr>{rpr}<w:vertAlign w:val="superscript"/></w:rPr>'
                f'<w:footnoteReference w:id="{i+2}"/></w:r></w:p>')
    body = "".join(fn_para(i) for i in range(n)) + sectpr()
    # footnotes.xml: separator (-1), continuationSeparator (0), then real notes
    fnrpr = f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/><w:sz w:val="18"/>'
    def note(i):
        return (f'<w:footnote w:id="{i+2}"><w:p><w:pPr><w:rPr>{fnrpr}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{fnrpr}<w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
                f'<w:r><w:rPr>{fnrpr}</w:rPr><w:t xml:space="preserve">注記{i+1}：本条の適用に関する補足説明であって、実務上の取扱いを示すものである。</w:t></w:r></w:p></w:footnote>')
    footnotes = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes {DOC_NS}>'
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
        + "".join(note(i) for i in range(n)) + '</w:footnotes>')
    write_docx(out("probefn_footnotes.docx"), doc(body),
               extra_parts={"word/footnotes.xml": footnotes},
               ct_extra=('<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',),
               rel_extra=('<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',))

def p_omml():
    """Display equations interleaved with body, near page boundaries."""
    rpr = f'<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="21"/>'
    def eq(k):
        # display fraction (a_k + b_k) / (c_k) as an oMathPara
        return (
            '<w:p><m:oMathPara><m:oMath>'
            '<m:f><m:fPr><m:ctrlPr><w:rPr/></m:ctrlPr></m:fPr>'
            f'<m:num><m:r><m:t>a+b+{k}</m:t></m:r></m:num>'
            f'<m:den><m:r><m:t>c-{k}</m:t></m:r></m:den></m:f>'
            '<m:r><m:t>=</m:t></m:r>'
            f'<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/><m:ctrlPr><w:rPr/></m:ctrlPr></m:naryPr>'
            f'<m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup>'
            f'<m:e><m:r><m:t>x_i</m:t></m:r></m:e></m:nary>'
            '</m:oMath></m:oMathPara></w:p>')
    # interleave: 3 body paras then an equation, repeated
    chunks = []
    for i in range(24):
        chunks.append(para(i + 1))
        if i % 3 == 2:
            chunks.append(eq(i))
    body = "".join(chunks) + sectpr()
    write_docx(out("probeomml_equations.docx"), doc(body))

def p_nested_table_split():
    """A table whose single cell holds a nested table + text, spanning pages."""
    rpr = f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/><w:sz w:val="21"/>'
    def cellp(txt):
        return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{rpr}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')
    inner = (
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
        + "".join(f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{cellp("内訳項目" + str(j+1) + "：" + SENT[:40])}</w:tc></w:tr>' for j in range(20))
        + '</w:tbl>' + cellp("（続き）"))
    outer = (
        '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>'
        '<w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="8" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{cellp("別表　明細")}{inner}</w:tc></w:tr>'
        '</w:tbl>')
    body = body_paras(10) + outer + body_paras(6) + sectpr()
    write_docx(out("probenest_nestedtable.docx"), doc(body))

PROBES = [p_vertical, p_twocol, p_threecol, p_b5, p_landscape, p_mixedgeom,
          p_grid_lac, p_grid_pitch480, p_yugothic, p_meiryo,
          p_linespacing_exact, p_linespacing_mult, p_nogrid_body,
          p_footnotes, p_omml, p_nested_table_split]

if __name__ == "__main__":
    os.makedirs(DOCX, exist_ok=True)
    for fn in PROBES:
        try:
            fn()
            print("ok  ", fn.__name__)
        except Exception as e:
            print("FAIL", fn.__name__, e)
