# -*- coding: utf-8 -*-
"""Adversarial structural probe batch 3 — section machinery, floats/wrap,
break runs, and reservation-estimate blind spots.

Batches 1-2 covered geometry/fonts/spacing/columns/tables/footnotes/ruby/
header-footer HEIGHT. Batch 3 targets what neither batch touched:
  - break RUNS (column break, mid-paragraph page break)
  - section machinery (continuous margin change, continuous 1->2->1 col on
    the SAME page, evenPage/oddPage starts, titlePg first-page header, gutter)
  - endnotes (footnotes' untested sibling)
  - floats with SIDE-wrap (anchored image wrapSquare, narrow floating table,
    VML textbox w10:wrap square) — the corpus only exercises wrap-below/none
  - wrap-content classes (long URLs in CJK, pure-Latin body, w:kern=2 mixed)
  - line-height stress (inline 24pt runs in a typed grid, exact < font clip)
  - reservation-estimate blind spots (header/footer with ONE WRAPPING
    paragraph — the height estimate is 1-line-per-para, width-independent)
  - pBdr boxed groups crossing a page, spacer-empty-paragraph runs,
    exact/atLeast trHeight rows, framePr positioned frames

All doc_ids start with "probex" so this batch gates independently of the
31 batch-1/2 probes ("probe2col" etc. do not match prefix "probex").

Run: python tools/metrics/_probe_gen3.py
"""
import os, sys, struct, zlib
sys.path.insert(0, os.path.dirname(__file__))
import _probe_gen as pg

MINCHO = pg.MINCHO
esc = pg.esc
SENT = pg.SENT

def out(n): return pg.out(n)

def rpr(sz="21", extra=""):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{sz}"/>{extra}')

def P(txt, jc="both", ppr="", sz="21", rp="", runs_after=""):
    r = rpr(sz, rp)
    return (f'<w:p><w:pPr><w:jc w:val="{jc}"/>{ppr}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r>'
            f'{runs_after}</w:p>')

def cond(n, **kw):
    return P(f"第{n}条　{SENT}", **kw)

def conds(a, b, **kw):
    return "".join(cond(i, **kw) for i in range(a, b + 1))

def brk_para(sp):
    return f'<w:p><w:pPr>{sp}</w:pPr></w:p>'

# ---- helpers: minimal PNG + picture XML ------------------------------------

def png_bytes(w=8, h=8, rgb=(176, 176, 216)):
    def chunk(t, d):
        return (struct.pack(">I", len(d)) + t + d
                + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF))
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))

A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
PIC_NS = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'

def _pic_xml(pid, cx, cy, rid):
    return (f'<a:graphic {A_NS}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic {PIC_NS}>'
            f'<pic:nvPicPr><pic:cNvPr id="{pid}" name="Pic{pid}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            f'</pic:pic></a:graphicData></a:graphic>')

def inline_img(pid, cx, cy, rid):
    return (f'<w:drawing><wp:inline {WP_NS} distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{pid}" name="Pic{pid}"/><wp:cNvGraphicFramePr/>'
            + _pic_xml(pid, cx, cy, rid) + '</wp:inline></w:drawing>')

def anchor_img(pid, cx, cy, rid):
    """Right-aligned floating picture with SQUARE wrap (text flows beside)."""
    return (f'<w:drawing><wp:anchor {WP_NS} distT="0" distB="0" distL="114300" distR="114300" '
            'simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:wrapSquare wrapText="bothSides"/>'
            f'<wp:docPr id="{pid}" name="Pic{pid}"/><wp:cNvGraphicFramePr/>'
            + _pic_xml(pid, cx, cy, rid) + '</wp:anchor></w:drawing>')

PNG_PART = {"word/media/image1.png": png_bytes()}
PNG_CT = ('<Default Extension="png" ContentType="image/png"/>',)
PNG_REL = ('<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>',)

# ---- 1. column-break runs in a 2-col section --------------------------------
def p_colbrk():
    parts = []
    for i in range(1, 41):
        if i % 8 == 0:
            parts.append(P(f"第{i}条　{SENT}",
                           runs_after='<w:r><w:br w:type="column"/></w:r>'))
        else:
            parts.append(cond(i))
    body = "".join(parts) + pg.sectpr(cols='<w:cols w:num="2" w:space="425"/>')
    pg.write_docx(out("probexcolbrk_columnbreak.docx"), pg.doc(body))

# ---- 2. mid-paragraph manual page break -------------------------------------
def p_pgbrk_mid():
    parts = []
    for i in range(1, 26):
        if i in (4, 11, 18):
            r = rpr()
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT[:50])}</w:t></w:r>'
                '<w:r><w:br w:type="page"/></w:r>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">（改ページ後の続き）{esc(SENT[50:])}</w:t></w:r></w:p>')
        else:
            parts.append(cond(i))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probexpgbrkmid_midparapagebreak.docx"), pg.doc(body))

# ---- 3. continuous section break with DIFFERENT L/R margins ------------------
def p_cont_margins():
    wide = '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>'
    narrow = '<w:pgMar w:top="1418" w:right="2835" w:bottom="1418" w:left="2835" w:header="851" w:footer="992" w:gutter="0"/>'
    s1 = conds(1, 20) + brk_para(pg.sectpr(mar=wide))
    s2 = conds(21, 45) + pg.sectpr(mar=narrow,
                                   sect_type='<w:type w:val="continuous"/>')
    pg.write_docx(out("probexmargins_contmarginchange.docx"), pg.doc(s1 + s2))

# ---- 4. continuous 1col -> 2col -> 1col on the SAME page ---------------------
def p_cont_2col():
    s1 = conds(1, 8) + brk_para(pg.sectpr(cols=''))
    s2 = conds(9, 44) + brk_para(pg.sectpr(
        cols='<w:cols w:num="2" w:space="425"/>',
        sect_type='<w:type w:val="continuous"/>'))
    s3 = conds(45, 52) + pg.sectpr(cols='',
                                   sect_type='<w:type w:val="continuous"/>')
    pg.write_docx(out("probexcont2col_contcolumns.docx"), pg.doc(s1 + s2 + s3))

# ---- 5. evenPage / oddPage section starts ------------------------------------
def p_evenodd():
    s1 = conds(1, 12) + brk_para(pg.sectpr())
    s2 = conds(13, 24) + brk_para(pg.sectpr(sect_type='<w:type w:val="evenPage"/>'))
    s3 = conds(25, 36) + pg.sectpr(sect_type='<w:type w:val="oddPage"/>')
    pg.write_docx(out("probexevenodd_evenoddsections.docx"), pg.doc(s1 + s2 + s3))

# ---- 5b. evenPage/oddPage with FORCED blank-page skips -----------------------
# v1 (probexevenodd) passed by PARITY LUCK: every section happened to end so
# the next natural page already had the required parity (no skip ever fired).
# v2 sizes the sections so BOTH skips must fire:
#   sec1 ends p2 -> natural next p3 (odd) -> evenPage skips to p4
#   sec2 p4..p5  -> natural next p6 (even) -> oddPage skips to p7
def p_evenodd2():
    s1 = conds(1, 16) + brk_para(pg.sectpr())
    s2 = conds(17, 32) + brk_para(pg.sectpr(sect_type='<w:type w:val="evenPage"/>'))
    s3 = conds(33, 44) + pg.sectpr(sect_type='<w:type w:val="oddPage"/>')
    pg.write_docx(out("probexeo2_evenoddskip.docx"), pg.doc(s1 + s2 + s3))

# ---- 6. titlePg: tall FIRST-page header, short default header ----------------
def p_titlepg():
    first = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<w:hdr {pg.DOC_NS}>'
             + "".join(P(f"表紙ヘッダ行{i+1}：本文書は試験用の見本である。") for i in range(6))
             + '</w:hdr>')
    dflt = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:hdr {pg.DOC_NS}>' + P("通常ヘッダ：見本") + '</w:hdr>')
    refs = ('<w:headerReference w:type="first" r:id="rId10"/>'
            '<w:headerReference w:type="default" r:id="rId11"/><w:titlePg/>')
    body = conds(1, 55) + pg.sectpr(sect_type=refs)
    pg.write_docx(out("probextitlepg_firstpageheader.docx"), pg.doc(body),
                  extra_parts={"word/header1.xml": first, "word/header2.xml": dflt},
                  ct_extra=('<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
                            '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'),
                  rel_extra=('<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
                             '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>'))

# ---- 7. gutter margin (binding edge) -----------------------------------------
def p_gutter():
    mar = '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="992" w:gutter="720"/>'
    body = conds(1, 55) + pg.sectpr(mar=mar)
    pg.write_docx(out("probexgutter_guttermargin.docx"), pg.doc(body))

# ---- 8. endnotes (sibling of the footnote breaker) ---------------------------
def p_endnotes():
    n = 30
    r = rpr()
    sup = rpr(extra='<w:vertAlign w:val="superscript"/>')
    def en_para(i):
        return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i+1}条　{esc(SENT)}</w:t></w:r>'
                f'<w:r><w:rPr>{sup}</w:rPr><w:endnoteReference w:id="{i+2}"/></w:r></w:p>')
    body = "".join(en_para(i) for i in range(n)) + pg.sectpr()
    enrpr = rpr("18")
    def note(i):
        return (f'<w:endnote w:id="{i+2}"><w:p><w:pPr><w:rPr>{enrpr}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{enrpr}<w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>'
                f'<w:r><w:rPr>{enrpr}</w:rPr><w:t xml:space="preserve">文末注{i+1}：本条の適用に関する補足説明であって、実務上の取扱いを示すものである。</w:t></w:r></w:p></w:endnote>')
    endnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:endnotes {pg.DOC_NS}>'
                '<w:endnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:endnote>'
                '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>'
                + "".join(note(i) for i in range(n)) + '</w:endnotes>')
    pg.write_docx(out("probexendnote_endnotes.docx"), pg.doc(body),
                  extra_parts={"word/endnotes.xml": endnotes},
                  ct_extra=('<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>',),
                  rel_extra=('<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>',))

# ---- 9. floating picture with wrapSquare (text flows BESIDE) ------------------
def p_img_float():
    parts = []
    for i in range(1, 51):
        if i in (5, 25):
            r = rpr()
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r>{anchor_img(i, 1800000, 1440000, "rId20")}</w:r>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT)}</w:t></w:r></w:p>')
        else:
            parts.append(cond(i))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probeximgfloat_wrapsquarepic.docx"), pg.doc(body),
                  extra_parts=dict(PNG_PART), ct_extra=PNG_CT, rel_extra=PNG_REL)

# ---- 10. inline pictures near page bottoms ------------------------------------
def p_img_inline():
    parts = []
    pid = 100
    for i in range(1, 41):
        parts.append(cond(i))
        if i % 4 == 0:
            pid += 1
            parts.append(f'<w:p><w:r>{inline_img(pid, 1440000, 1080000, "rId20")}</w:r></w:p>')
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probeximginline_inlinepics.docx"), pg.doc(body),
                  extra_parts=dict(PNG_PART), ct_extra=PNG_CT, rel_extra=PNG_REL)

# ---- 11. NARROW floating table (text flows beside) ----------------------------
def p_float_tbl():
    def cellp(txt):
        r = rpr()
        return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')
    def ftbl(tag):
        rows = "".join(
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="3400" w:type="dxa"/></w:tcPr>{cellp(f"{tag}項目{j+1}：数値{(j+1)*7}")}</w:tc></w:tr>'
            for j in range(8))
        return ('<w:tbl><w:tblPr>'
                '<w:tblpPr w:leftFromText="142" w:rightFromText="142" w:vertAnchor="text" '
                'w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="1"/>'
                '<w:tblW w:w="3400" w:type="dxa"/>' + pg2_borders() + '</w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="3400"/></w:tblGrid>' + rows + '</w:tbl>')
    body = (conds(1, 5) + ftbl("甲") + conds(6, 28) + ftbl("乙")
            + conds(29, 48) + pg.sectpr())
    pg.write_docx(out("probexfloattbl_narrowfloattable.docx"), pg.doc(body))

def pg2_borders():
    return ('<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>')

# ---- 12. VML textbox with square wrap ------------------------------------------
def p_vml_txbx():
    inner = "".join(P(f"囲み注記{j+1}：試験用の見本である。", jc="left", sz="18") for j in range(3))
    pict = ('<w:pict>'
            '<v:rect xmlns:v="urn:schemas-microsoft-com:vml" id="TB1" '
            'style="position:absolute;margin-left:0;margin-top:2pt;width:160pt;height:80pt;z-index:2;'
            'mso-position-horizontal:right;mso-position-horizontal-relative:margin;'
            'mso-position-vertical-relative:text;mso-wrap-distance-left:9pt;mso-wrap-distance-right:9pt" '
            'fillcolor="#eeeeee" strokecolor="black" strokeweight="1pt">'
            f'<v:textbox><w:txbxContent>{inner}</w:txbxContent></v:textbox></v:rect>'
            '<w10:wrap xmlns:w10="urn:schemas-microsoft-com:office:word" type="square" anchorx="margin"/>'
            '</w:pict>')
    r = rpr()
    tb_para = (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
               f'<w:r>{pict}</w:r>'
               f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第6条　{esc(SENT)}</w:t></w:r></w:p>')
    body = conds(1, 5) + tb_para + conds(7, 48) + pg.sectpr()
    pg.write_docx(out("probextxbx_vmltextboxwrap.docx"), pg.doc(body))

# ---- 13. long URLs inside CJK body ---------------------------------------------
def p_urls():
    def upara(i):
        url = (f"https://www.example.co.jp/houki/2026/download/annual-report_v{i}.html"
               f"?id={1000+i}&lang=ja&mode=print")
        return P(f"第{i}条　本条に関する資料は {url} に掲載するものとし、"
                 + SENT[:60])
    body = "".join(upara(i) for i in range(1, 46)) + pg.sectpr()
    pg.write_docx(out("probexurl_longurls.docx"), pg.doc(body))

# ---- 14. pure-Latin justified body (no grid) ------------------------------------
LATIN_SENT = ("The provisions of this Agreement shall be interpreted in good faith "
              "and applied consistently across all departments, and any matter arising "
              "in connection with the implementation of these provisions shall be "
              "resolved promptly, fairly, and without undue delay by the parties concerned.")

def p_latin():
    def lp(i):
        r = f'<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>'
        return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">Article {i}. {LATIN_SENT}</w:t></w:r></w:p>')
    body = "".join(lp(i) for i in range(1, 51)) + pg.sectpr(grid='')
    pg.write_docx(out("probexlatin_latinbody.docx"), pg.doc(body),
                  font="Calibri", sz="22", cpunct=False)

# ---- 15. w:kern=2 in docDefaults, kern-pair-heavy mixed body ---------------------
def p_kern():
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              f'<w:rFonts w:ascii="Times New Roman" w:eastAsia="{MINCHO}" w:hAnsi="Times New Roman"/>'
              '<w:kern w:val="2"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
              '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>')
    def kp(i):
        r = (f'<w:rFonts w:ascii="Times New Roman" w:eastAsia="{MINCHO}" w:hAnsi="Times New Roman"/>'
             '<w:kern w:val="2"/><w:sz w:val="21"/>')
        return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　AVATAR Type-V Wave PLAN および To/Ya 型の適用については、{esc(SENT[:80])}</w:t></w:r></w:p>')
    body = "".join(kp(i) for i in range(1, 51)) + pg.sectpr()
    pg.write_docx(out("probexkern_kernpairs.docx"), pg.doc(body),
                  extra_parts={"word/styles.xml": styles})

# ---- 16. inline 24pt runs inside 10.5pt typed-grid body --------------------------
def p_bigrun():
    parts = []
    for i in range(1, 46):
        if i % 3 == 0:
            r = rpr()
            big = rpr("48")
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT[:40])}</w:t></w:r>'
                f'<w:r><w:rPr>{big}</w:rPr><w:t xml:space="preserve">重要</w:t></w:r>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(SENT[40:])}</w:t></w:r></w:p>')
        else:
            parts.append(cond(i))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probexbigrun_inlinebigruns.docx"), pg.doc(body))

# ---- 17. exact line spacing SMALLER than the font (clipped lines) ----------------
def p_exact_clip():
    ppr = '<w:spacing w:line="160" w:lineRule="exact"/>'
    body = conds(1, 80, ppr=ppr) + pg.sectpr()
    pg.write_docx(out("probexexactclip_exactclip.docx"), pg.doc(body))

# ---- 18. identical-pBdr boxed groups crossing page boundaries --------------------
def p_pbdr_groups():
    pbdr = ('<w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="auto"/>'
            '<w:left w:val="single" w:sz="8" w:space="4" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="8" w:space="1" w:color="auto"/>'
            '<w:right w:val="single" w:sz="8" w:space="4" w:color="auto"/></w:pBdr>')
    parts = []
    k = 1
    for g in range(14):
        for j in range(3):
            parts.append(P(f"第{k}条　{SENT}", ppr=pbdr))
            k += 1
        parts.append(P(f"（第{g+1}群の解説）" + SENT[:60]))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probexpbdr_borderedgroups.docx"), pg.doc(body))

# ---- 19. spacer-heavy: text + runs of empty paragraphs ---------------------------
def p_empty_spam():
    r = rpr()
    empty = f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr></w:p>'
    parts = []
    for i in range(1, 27):
        parts.append(cond(i))
        parts.append(empty * 4)
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probexempty_spacerempties.docx"), pg.doc(body))

# ---- 20. exact / atLeast trHeight rows spanning pages ----------------------------
def p_trheight():
    def cellp(txt):
        r = rpr()
        return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')
    rows = []
    for i in range(36):
        if i % 2 == 0:
            trpr = '<w:trPr><w:trHeight w:val="400" w:hRule="exact"/></w:trPr>'
            content = cellp(f"第{i+1}項（exact 400tw、2行分の内容を切詰め）：{SENT[:50]}")
        else:
            trpr = '<w:trPr><w:trHeight w:val="800" w:hRule="atLeast"/></w:trPr>'
            content = cellp(f"第{i+1}項（atLeast 800tw、内容1行）")
        rows.append(f'<w:tr>{trpr}<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{content}</w:tc></w:tr>')
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + pg2_borders() + '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' + "".join(rows) + '</w:tbl>')
    body = conds(1, 3) + tbl + conds(4, 6) + pg.sectpr()
    pg.write_docx(out("probextrheight_exactatleastrows.docx"), pg.doc(body))

# ---- 21/22. header / footer with ONE WRAPPING paragraph --------------------------
# (the header/footer height ESTIMATE is 1-line-per-para, width-independent —
#  a single paragraph that wraps to ~5 lines is the estimate's blind spot)
def _wrap_hdr_ftr(name, which):
    long_para = P("本文書の管理区分：" + SENT + SENT[:70], jc="left")
    hdr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:hdr {pg.DOC_NS}>' + long_para + '</w:hdr>')
    ftr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:ftr {pg.DOC_NS}>' + long_para + '</w:ftr>')
    if which == "header":
        ref = '<w:headerReference w:type="default" r:id="rId10"/>'
        part = {"word/header1.xml": hdr}
        ct = ('<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',)
        rel = ('<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',)
    else:
        ref = '<w:footerReference w:type="default" r:id="rId10"/>'
        part = {"word/footer1.xml": ftr}
        ct = ('<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',)
        rel = ('<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',)
    body = conds(1, 55) + pg.sectpr(sect_type=ref)
    pg.write_docx(out(name), pg.doc(body), extra_parts=part, ct_extra=ct, rel_extra=rel)

def p_hdr_wrap(): _wrap_hdr_ftr("probexhdrwrap_wrappingheader.docx", "header")
def p_ftr_wrap(): _wrap_hdr_ftr("probexftrwrap_wrappingfooter.docx", "footer")

# ---- 23. framePr positioned text frames -------------------------------------------
def p_frames():
    parts = []
    r = rpr()
    for i in range(1, 46):
        parts.append(cond(i))
        if i % 10 == 0:
            fp = ('<w:framePr w:w="2600" w:h="1100" w:hRule="atLeast" w:hSpace="141" '
                  'w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="right" w:y="1"/>')
            # framePr must precede jc in pPr (CT_PPr sequence)
            parts.append(
                f'<w:p><w:pPr>{fp}<w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">（枠内注記{i//10}）別紙を参照のこと。</w:t></w:r></w:p>')
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probexframe_textframes.docx"), pg.doc(body))

PROBES = [p_colbrk, p_pgbrk_mid, p_cont_margins, p_cont_2col, p_evenodd,
          p_evenodd2, p_titlepg, p_gutter, p_endnotes, p_img_float,
          p_img_inline, p_float_tbl, p_vml_txbx, p_urls, p_latin, p_kern,
          p_bigrun, p_exact_clip, p_pbdr_groups, p_empty_spam, p_trheight,
          p_hdr_wrap, p_ftr_wrap, p_frames]

if __name__ == "__main__":
    for fn in PROBES:
        try:
            fn(); print("ok  ", fn.__name__)
        except Exception as e:
            print("FAIL", fn.__name__, e)
