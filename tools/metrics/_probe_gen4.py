# -*- coding: utf-8 -*-
"""Adversarial structural probe batch 4 — content-DROP wrappers, wrap-rule
TOGGLES, header machinery variants, and remaining structure classes.

Batches 1-3 (55 probes) covered geometry/fonts/spacing/columns/tables/notes/
floats/side-wrap/sections/estimates. Batch 4 targets what none touched:
  - content-drop wrappers: w:sdt (content controls — ubiquitous in modern
    forms), w:smartTag (legacy wrapper) — if sdtContent/smartTag children are
    unparsed the text vanishes (the probexendnote MASKED-breaker class; the
    S724 LOW_MATCH net catches it honestly now)
  - wrap-rule toggles Oxi's tuned defaults never see: kinsoku=0, wordWrap=0
    (mid-word Latin breaks), overflowPunct=0 (no ぶら下げ), autoSpaceDE/DN=0,
    settings autoHyphenation (Latin)
  - header machinery: TABLE inside a header (the 1-line-per-para estimate has
    no table arm), per-SECTION different-height headers, evenAndOddHeaders
  - floats: wrapTopAndBottom (full-width flow reservation — batch 3 did only
    wrapSquare), INLINE VML textbox (in-line box, not floating)
  - tables: gridSpan rows crossing pages, tbRlV rotated cells (tall rows)
  - grid: POSITIVE docGrid charSpace, continuous section with a DIFFERENT
    linePitch mid-page (per-section grid on a merged page — S560 family)
  - misc: consecutive manual page breaks (blank pages), negative/extreme
    indents, non-BMP/surrogate/combining chars

All doc_ids start with "probez" (NOT "probey" — "probeyug" from batch 1 would
collide with a "probey" prefix filter).

Run: python tools/metrics/_probe_gen4.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import _probe_gen as pg
import _probe_gen3 as g3

MINCHO = pg.MINCHO
esc = pg.esc
SENT = pg.SENT
rpr, P, cond, conds, brk_para = g3.rpr, g3.P, g3.cond, g3.conds, g3.brk_para

def out(n): return pg.out(n)

HDR_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
HDR_RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'

# ---- 1. content controls (w:sdt), block + run level -------------------------
def p_sdt():
    r = rpr()
    parts = []
    for i in range(1, 37):
        body_p = cond(i)
        if i % 3 == 0:
            # block-level SDT wrapping the whole paragraph
            parts.append(f'<w:sdt><w:sdtPr><w:id w:val="{i}"/></w:sdtPr>'
                         f'<w:sdtContent>{body_p}</w:sdtContent></w:sdt>')
        elif i % 3 == 1:
            # run-level SDT wrapping the middle of the paragraph
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT[:30])}</w:t></w:r>'
                f'<w:sdt><w:sdtPr><w:id w:val="{1000+i}"/></w:sdtPr><w:sdtContent>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(SENT[30:60])}</w:t></w:r>'
                f'</w:sdtContent></w:sdt>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(SENT[60:])}</w:t></w:r></w:p>')
        else:
            parts.append(body_p)
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probezsdt_contentcontrols.docx"), pg.doc(body))

# ---- 2. legacy smartTag run wrapper ------------------------------------------
def p_smarttag():
    r = rpr()
    parts = []
    for i in range(1, 41):
        if i % 2 == 0:
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT[:25])}</w:t></w:r>'
                f'<w:smartTag w:uri="urn:probe-schema" w:element="term">'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(SENT[25:55])}</w:t></w:r>'
                f'</w:smartTag>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(SENT[55:])}</w:t></w:r></w:p>')
        else:
            parts.append(cond(i))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probezsmarttag_smarttagruns.docx"), pg.doc(body))

# ---- 3. anchored picture wrapTopAndBottom (full-width flow reservation) ------
def anchor_img_tb(pid, cx, cy, rid):
    return (f'<w:drawing><wp:anchor {g3.WP_NS} distT="0" distB="0" distL="114300" distR="114300" '
            'simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:wrapTopAndBottom/>'
            f'<wp:docPr id="{pid}" name="Pic{pid}"/><wp:cNvGraphicFramePr/>'
            + g3._pic_xml(pid, cx, cy, rid) + '</wp:anchor></w:drawing>')

def p_wraptb():
    r = rpr()
    parts = []
    for i in range(1, 46):
        if i in (5, 25):
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r>{anchor_img_tb(i, 2600000, 1600000, "rId20")}</w:r>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT)}</w:t></w:r></w:p>')
        else:
            parts.append(cond(i))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probezwraptb_wraptopbottom.docx"), pg.doc(body),
                  extra_parts=dict(g3.PNG_PART), ct_extra=g3.PNG_CT, rel_extra=g3.PNG_REL)

# ---- 4. TABLE inside the header (estimate has no table arm) -------------------
def p_hdr_table():
    def cellp(txt):
        r = rpr()
        return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + g3.pg2_borders() + '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>'
           + "".join(
               f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>{cellp(lab)}</w:tc>'
               f'<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>{cellp(val)}</w:tc></w:tr>'
               for lab, val in [("文書番号", "PRB-2026-004"), ("管理区分", "試験用見本"),
                                ("改訂日", "2026年7月3日")])
           + '</w:tbl>')
    hdr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:hdr {pg.DOC_NS}>' + tbl + P("（ヘッダ末尾行）", jc="left") + '</w:hdr>')
    ref = '<w:headerReference w:type="default" r:id="rId10"/>'
    body = conds(1, 55) + pg.sectpr(sect_type=ref)
    pg.write_docx(out("probezhdrtbl_tableinheader.docx"), pg.doc(body),
                  extra_parts={"word/header1.xml": hdr},
                  ct_extra=(f'<Override PartName="/word/header1.xml" ContentType="{HDR_CT}"/>',),
                  rel_extra=(f'<Relationship Id="rId10" Type="{HDR_RT}" Target="header1.xml"/>',))

# ---- 5. per-SECTION headers with different heights (nextPage) -----------------
def p_sect_headers():
    small = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             f'<w:hdr {pg.DOC_NS}>' + P("第一部ヘッダ：見本", jc="left") + '</w:hdr>')
    tall = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:hdr {pg.DOC_NS}>'
            + "".join(P(f"第二部ヘッダ行{i+1}：本文書は試験用の見本である。", jc="left") for i in range(6))
            + '</w:hdr>')
    ref1 = '<w:headerReference w:type="default" r:id="rId10"/>'
    ref2 = '<w:headerReference w:type="default" r:id="rId11"/>'
    s1 = conds(1, 16) + brk_para(pg.sectpr(sect_type=ref1))
    s2 = conds(17, 44) + pg.sectpr(sect_type=ref2)
    pg.write_docx(out("probezsecthdr_persectionheaders.docx"), pg.doc(s1 + s2),
                  extra_parts={"word/header1.xml": small, "word/header2.xml": tall},
                  ct_extra=(f'<Override PartName="/word/header1.xml" ContentType="{HDR_CT}"/>',
                            f'<Override PartName="/word/header2.xml" ContentType="{HDR_CT}"/>'),
                  rel_extra=(f'<Relationship Id="rId10" Type="{HDR_RT}" Target="header1.xml"/>',
                             f'<Relationship Id="rId11" Type="{HDR_RT}" Target="header2.xml"/>'))

# ---- 6. evenAndOddHeaders with different heights ------------------------------
def p_evenodd_headers():
    odd = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:hdr {pg.DOC_NS}>'
           + "".join(P(f"奇数頁ヘッダ行{i+1}：本文書は試験用の見本である。", jc="left") for i in range(5))
           + '</w:hdr>')
    even = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:hdr {pg.DOC_NS}>' + P("偶数頁ヘッダ：見本", jc="left") + '</w:hdr>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:evenAndOddHeaders/>'
                '<w:characterSpacingControl w:val="compressPunctuation"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    refs = ('<w:headerReference w:type="default" r:id="rId10"/>'
            '<w:headerReference w:type="even" r:id="rId11"/>')
    body = conds(1, 55) + pg.sectpr(sect_type=refs)
    pg.write_docx(out("probezevenhdr_evenoddheaders.docx"), pg.doc(body),
                  extra_parts={"word/header1.xml": odd, "word/header2.xml": even,
                               "word/settings.xml": settings},
                  ct_extra=(f'<Override PartName="/word/header1.xml" ContentType="{HDR_CT}"/>',
                            f'<Override PartName="/word/header2.xml" ContentType="{HDR_CT}"/>'),
                  rel_extra=(f'<Relationship Id="rId10" Type="{HDR_RT}" Target="header1.xml"/>',
                             f'<Relationship Id="rId11" Type="{HDR_RT}" Target="header2.xml"/>'))

# ---- 7-9. wrap-rule toggles: kinsoku=0 / wordWrap=0 / overflowPunct=0 ----------
def p_kinsoku_off():
    body = conds(1, 55, ppr='<w:kinsoku w:val="0"/>') + pg.sectpr()
    pg.write_docx(out("probezkinsoku_kinsokuoff.docx"), pg.doc(body))

LONGWORDS = ("Notwithstanding intergovernmental responsibilities and "
             "implementation-specific administrativeprocedures ")

def p_wordwrap_off():
    def wp(i):
        return P(f"第{i}条　{LONGWORDS}に関しては、{SENT[:60]}",
                 ppr='<w:wordWrap w:val="0"/>')
    body = "".join(wp(i) for i in range(1, 51)) + pg.sectpr()
    pg.write_docx(out("probezwordwrap_wordwrapoff.docx"), pg.doc(body))

def p_overflowpunct_off():
    body = conds(1, 55, ppr='<w:overflowPunct w:val="0"/>') + pg.sectpr()
    pg.write_docx(out("probezovflpunct_nohangingpunct.docx"), pg.doc(body))

# ---- 10. autoSpaceDE/DN = 0 (CJK-Latin/digit autospace off) --------------------
def p_autospace_off():
    def ap(i):
        return P(f"第{i}条　Word2026版のRevision4に基づきChapter{i}の規定を適用し、"
                 f"限度は{i * 13}時間かつ{i * 7}日とする。{SENT[:50]}",
                 ppr='<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>')
    body = "".join(ap(i) for i in range(1, 51)) + pg.sectpr()
    pg.write_docx(out("probezautospace_autospaceoff.docx"), pg.doc(body))

# ---- 11. settings autoHyphenation + Latin body ---------------------------------
def p_hyphenation():
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:autoHyphenation/><w:hyphenationZone w:val="360"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    def lp(i):
        r = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/>'
        return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">Article {i}. '
                f'{g3.LATIN_SENT} Notwithstanding the responsibilities, implementation '
                f'and administration shall remain uninterrupted.</w:t></w:r></w:p>')
    body = "".join(lp(i) for i in range(1, 46)) + pg.sectpr(grid='')
    pg.write_docx(out("probezhyph_autohyphenation.docx"), pg.doc(body),
                  font="Calibri", sz="22",
                  extra_parts={"word/settings.xml": settings})

# ---- 12. gridSpan rows crossing pages ------------------------------------------
def p_gridspan():
    def cellp(txt):
        r = rpr()
        return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')
    rows = []
    for i in range(40):
        if i % 3 == 0:
            rows.append('<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/>'
                        f'<w:gridSpan w:val="2"/></w:tcPr>{cellp(f"第{i+1}項（結合）：{SENT[:70]}")}</w:tc></w:tr>')
        else:
            rows.append('<w:tr>'
                        f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{cellp(f"第{i+1}項")}</w:tc>'
                        f'<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>{cellp(SENT[:60])}</w:tc>'
                        '</w:tr>')
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + g3.pg2_borders() + '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6000"/></w:tblGrid>'
           + "".join(rows) + '</w:tbl>')
    body = conds(1, 3) + tbl + conds(4, 6) + pg.sectpr()
    pg.write_docx(out("probezgridspan_mergedcols.docx"), pg.doc(body))

# ---- 13. tbRlV rotated cells (tall rows crossing pages) ------------------------
def p_tbrlv():
    def vcell(txt):
        r = rpr()
        return (f'<w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/>'
                '<w:textDirection w:val="tbRlV"/><w:vAlign w:val="center"/></w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p></w:tc>')
    def hcell(txt):
        r = rpr()
        return (f'<w:tc><w:tcPr><w:tcW w:w="7800" w:type="dxa"/></w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p></w:tc>')
    rows = "".join(
        f'<w:tr>{vcell(f"区分{i+1}　" + SENT[:18])}{hcell(f"第{i+1}項：" + SENT)}</w:tr>'
        for i in range(12))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + g3.pg2_borders() + '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="7800"/></w:tblGrid>'
           + rows + '</w:tbl>')
    body = conds(1, 2) + tbl + conds(3, 4) + pg.sectpr()
    pg.write_docx(out("probeztbrlv_rotatedcells.docx"), pg.doc(body))

# ---- 14. consecutive manual page breaks (blank pages) --------------------------
def p_pgbrk_multi():
    brk = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    body = (conds(1, 6) + brk * 3 + conds(7, 12) + brk * 2 + conds(13, 18)
            + pg.sectpr())
    pg.write_docx(out("probezpgbrk3_consecutivebreaks.docx"), pg.doc(body))

# ---- 15. negative / extreme indents --------------------------------------------
def p_neg_indent():
    parts = []
    for i in range(1, 51):
        if i % 3 == 0:
            ppr = '<w:ind w:left="-567" w:right="0"/>'          # into left margin
        elif i % 3 == 1:
            ppr = '<w:ind w:left="1985" w:hanging="1985"/>'     # big hanging
        else:
            ppr = '<w:ind w:left="0" w:right="1701"/>'          # big right indent
        parts.append(P(f"第{i}条　{SENT}", ppr=ppr))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probeznegind_extremeindents.docx"), pg.doc(body))

# ---- 16. INLINE VML textbox (in-line box grows the line) -----------------------
def p_inline_txbx():
    r = rpr()
    inner = P("箱内：見本", jc="left", sz="18")
    def pict():
        return ('<w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" '
                'style="width:120pt;height:44pt" fillcolor="#f0f0f0" '
                f'strokecolor="black" strokeweight="0.5pt"><v:textbox><w:txbxContent>{inner}'
                '</w:txbxContent></v:textbox></v:rect></w:pict>')
    parts = []
    for i in range(1, 41):
        if i % 8 == 0:
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT[:30])}</w:t></w:r>'
                f'<w:r>{pict()}</w:r>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(SENT[30:])}</w:t></w:r></w:p>')
        else:
            parts.append(cond(i))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probezinlinetb_inlinetextbox.docx"), pg.doc(body))

# ---- 17. POSITIVE docGrid charSpace (wider char pitch) --------------------------
def p_charspace_pos():
    grid = '<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="1966"/>'
    body = conds(1, 55) + pg.sectpr(grid=grid)
    pg.write_docx(out("probezcharsppos_positivecharspace.docx"), pg.doc(body))

# ---- 18. continuous section with a DIFFERENT linePitch mid-page ----------------
def p_cont_grid():
    s1 = conds(1, 14) + brk_para(pg.sectpr(
        grid='<w:docGrid w:type="lines" w:linePitch="360"/>'))
    s2 = conds(15, 44) + pg.sectpr(
        grid='<w:docGrid w:type="lines" w:linePitch="480"/>',
        sect_type='<w:type w:val="continuous"/>')
    pg.write_docx(out("probezcontgrid_gridpitchchange.docx"), pg.doc(s1 + s2))

# ---- 19. non-BMP / surrogate / combining chars ----------------------------------
def p_exotic_chars():
    exotic = "𠮟責、𩸽の𡈽産、㐂寿、髙﨑、あ゙ゔ、㊤㊥㊦、Ⅷ章"
    def xp(i):
        return P(f"第{i}条　{exotic}に関する取扱いは、{SENT[:60]}")
    body = "".join(xp(i) for i in range(1, 46)) + pg.sectpr()
    pg.write_docx(out("probezexotic_nonbmpchars.docx"), pg.doc(body))

PROBES = [p_sdt, p_smarttag, p_wraptb, p_hdr_table, p_sect_headers,
          p_evenodd_headers, p_kinsoku_off, p_wordwrap_off,
          p_overflowpunct_off, p_autospace_off, p_hyphenation, p_gridspan,
          p_tbrlv, p_pgbrk_multi, p_neg_indent, p_inline_txbx,
          p_charspace_pos, p_cont_grid, p_exotic_chars]

if __name__ == "__main__":
    for fn in PROBES:
        try:
            fn(); print("ok  ", fn.__name__)
        except Exception as e:
            print("FAIL", fn.__name__, e)
