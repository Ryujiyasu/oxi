# -*- coding: utf-8 -*-
"""Adversarial structural probe batch 2 — feature INTERACTIONS + under-tested
content features. Batch 1 showed single-feature geometry/fonts/spacing/OMML pass;
the breakers were flow/reservation paths (columns, footnotes, vertical, table
split). Batch 2 stresses interactions of those + content features.

Run: python tools/metrics/_probe_gen2.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import _probe_gen as pg

MINCHO = pg.MINCHO
esc = pg.esc
SENT = pg.SENT

def out(n): return pg.out(n)

def rpr(sz="21", extra=""):
    return f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/><w:sz w:val="{sz}"/>{extra}'

def P(txt, jc="both", ppr="", sz="21", rp=""):
    r = rpr(sz, rp)
    return (f'<w:p><w:pPr><w:jc w:val="{jc}"/>{ppr}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

def cond(n, **kw):
    return P(f"第{n}条　{SENT}", **kw)

def conds(a, b, **kw):
    return "".join(cond(i, **kw) for i in range(a, b + 1))

# ---- 1. mixed column counts across sections (1 -> 2 -> 1) ------------------
def p_cols_mixed():
    def brk(sp): return f'<w:p><w:pPr>{sp}</w:pPr></w:p>'
    s1 = conds(1, 14) + brk(pg.sectpr(cols=''))                       # 1-col
    s2 = conds(15, 44) + brk(pg.sectpr(cols='<w:cols w:num="2" w:space="425" w:sep="1"/>'))  # 2-col
    s3 = conds(45, 60) + pg.sectpr(cols='')                          # 1-col (final)
    pg.write_docx(out("probecolsmix_mixedcols.docx"), pg.doc(s1 + s2 + s3))

# ---- 2. unequal-width 2 columns -------------------------------------------
def p_cols_unequal():
    cols = ('<w:cols w:num="2" w:space="425" w:equalWidth="0">'
            '<w:col w:w="3200" w:space="425"/><w:col w:w="5445"/></w:cols>')
    body = conds(1, 60) + pg.sectpr(cols=cols)
    pg.write_docx(out("probecolsuneq_unequalcols.docx"), pg.doc(body))

# ---- 3. tall multi-line header (eats body area) ---------------------------
def _hdr_ftr_docx(name, which):
    hdr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:hdr {pg.DOC_NS}>' + "".join(P(f"ヘッダ行{i+1}：本文書は試験用の見本である。") for i in range(6)) + '</w:hdr>')
    ftr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:ftr {pg.DOC_NS}>' + "".join(P(f"フッタ行{i+1}：試験用の見本である。") for i in range(6)) + '</w:ftr>')
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
    # keep the default header/footer margins so the tall block must eat body area
    sect = pg.sectpr(sect_type=ref)
    body = conds(1, 55) + sect
    pg.write_docx(out(name), pg.doc(body), extra_parts=part, ct_extra=ct, rel_extra=rel)

def p_header_tall(): _hdr_ftr_docx("probehdrtall_tallheader.docx", "header")
def p_footer_tall(): _hdr_ftr_docx("probeftrtall_tallfooter.docx", "footer")

# ---- table helpers --------------------------------------------------------
def _cellp(txt, sz="21"):
    r = rpr(sz)
    return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

TBORD = ('<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>')

# ---- 4. long table with a REPEATING header row (tblHeader) -----------------
def p_table_header_repeat():
    header_row = ('<w:tr><w:trPr><w:tblHeader/></w:trPr>'
                  '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' + _cellp("項目") + '</w:tc>'
                  '<w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>' + _cellp("内容") + '</w:tc></w:tr>')
    rows = "".join(
        '<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>{_cellp("第" + str(i+1) + "項")}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>{_cellp(SENT)}</w:tc>'
        '</w:tr>' for i in range(50))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + TBORD + '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="7000"/></w:tblGrid>'
           + header_row + rows + '</w:tbl>')
    body = conds(1, 3) + tbl + conds(4, 5) + pg.sectpr()
    pg.write_docx(out("probethdr_tableheaderrepeat.docx"), pg.doc(body))

# ---- 5. table rows that CANNOT split (cantSplit) near page boundary --------
def p_table_cantsplit():
    def row(i):
        # tall multi-para cell that would straddle a page → cantSplit forces whole-row move
        cells = "".join(_cellp(f"第{i+1}項の{j+1}: " + SENT) for j in range(4))
        return ('<w:tr><w:trPr><w:cantSplit/></w:trPr>'
                f'<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{cells}</w:tc></w:tr>')
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + TBORD + '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
           + "".join(row(i) for i in range(12)) + '</w:tbl>')
    body = conds(1, 2) + tbl + pg.sectpr()
    pg.write_docx(out("probecantsplit_cantsplit.docx"), pg.doc(body))

# ---- 6. vertically merged cell spanning a page boundary --------------------
def p_table_vmerge():
    def row(i, first):
        vm = '<w:vMerge w:val="restart"/>' if first else '<w:vMerge/>'
        left = (f'<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/>{vm}</w:tcPr>'
                + (_cellp("大区分（縦結合）") if first else '<w:p/>') + '</w:tc>')
        right = f'<w:tc><w:tcPr><w:tcW w:w="6500" w:type="dxa"/></w:tcPr>{_cellp(f"第{i+1}項: " + SENT)}</w:tc>'
        return f'<w:tr>{left}{right}</w:tr>'
    # one merged block of 30 rows (spans multiple pages)
    rows = "".join(row(i, i == 0) for i in range(30))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + TBORD + '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="6500"/></w:tblGrid>'
           + rows + '</w:tbl>')
    body = conds(1, 2) + tbl + pg.sectpr()
    pg.write_docx(out("probevmerge_vmergespan.docx"), pg.doc(body))

# ---- 7. ruby (furigana) in body flow, multi-page --------------------------
def p_ruby_body():
    r = rpr()
    def ruby_run(base, rt):
        return ('<w:r><w:ruby>'
                '<w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="10"/><w:hpsRaise w:val="18"/><w:hpsBaseText w:val="21"/><w:lid w:val="ja-JP"/></w:rubyPr>'
                f'<w:rt><w:r><w:rPr><w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/><w:sz w:val="10"/></w:rPr><w:t>{esc(rt)}</w:t></w:r></w:rt>'
                f'<w:rubyBase><w:r><w:rPr>{r}</w:rPr><w:t>{esc(base)}</w:t></w:r></w:rubyBase>'
                '</w:ruby></w:r>')
    def para(i):
        return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i+1}条　</w:t></w:r>'
                + ruby_run("本項", "ほんこう")
                + f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">に定める</w:t></w:r>'
                + ruby_run("事項", "じこう")
                + f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">については、{esc(SENT)}</w:t></w:r></w:p>')
    body = "".join(para(i) for i in range(50)) + pg.sectpr()
    pg.write_docx(out("proberuby_rubybody.docx"), pg.doc(body))

# ---- 8. numbered multi-level list (numbering.xml), multi-page --------------
def p_list_numbered():
    numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 f'<w:numbering {pg.DOC_NS}>'
                 '<w:abstractNum w:abstractNumId="0">'
                 + "".join(
                     f'<w:lvl w:ilvl="{lv}"><w:start w:val="1"/>'
                     f'<w:numFmt w:val="decimal"/><w:lvlText w:val="%{lv+1}."/>'
                     f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{(lv+1)*480}" w:hanging="480"/></w:pPr></w:lvl>'
                     for lv in range(3))
                 + '</w:abstractNum>'
                 '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
    def item(i):
        lv = i % 3
        ppr = f'<w:numPr><w:ilvl w:val="{lv}"/><w:numId w:val="1"/></w:numPr>'
        return P(SENT, ppr=ppr)
    body = "".join(item(i) for i in range(70)) + pg.sectpr()
    pg.write_docx(out("problist_numberedlist.docx"), pg.doc(body),
                  extra_parts={"word/numbering.xml": numbering},
                  ct_extra=('<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',),
                  rel_extra=('<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',))

# ---- 9. page borders (pgBorders) — do they shrink body area? ---------------
def p_page_border():
    pgb = ('<w:pgBorders w:offsetFrom="page">'
           '<w:top w:val="single" w:sz="24" w:space="24" w:color="auto"/>'
           '<w:left w:val="single" w:sz="24" w:space="24" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="24" w:space="24" w:color="auto"/>'
           '<w:right w:val="single" w:sz="24" w:space="24" w:color="auto"/></w:pgBorders>')
    sect = f'<w:sectPr>{pgb}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'
    body = conds(1, 55) + sect
    pg.write_docx(out("probepgborder_pageborder.docx"), pg.doc(body))

# ---- 10. widowControl=1 body straddling page boundaries -------------------
def p_widow():
    # override the Normal widowControl=0 by setting it ON per-paragraph
    ppr = '<w:widowControl/>'
    body = conds(1, 55, ppr=ppr) + pg.sectpr()
    pg.write_docx(out("probewidow_widowcontrol.docx"), pg.doc(body))

# ---- 11. keepNext heading chains near page boundaries ---------------------
def p_keepnext():
    parts = []
    for i in range(40):
        # a heading kept with the next para
        parts.append(P(f"第{i+1}章　総則に関する見出し", jc="left", ppr='<w:keepNext/><w:keepLines/>', sz="24"))
        parts.append(P(SENT))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probekeepnext_keepnext.docx"), pg.doc(body))

# ---- 12. drop caps in body flow, multi-page -------------------------------
def p_dropcap():
    r = rpr()
    def block(i):
        # a dropCap frame paragraph + its body paragraph
        cap = (f'<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>'
               f'<w:rPr><w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/><w:sz w:val="21"/><w:position w:val="-4"/></w:rPr></w:pPr>'
               f'<w:r><w:rPr><w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/><w:sz w:val="21"/></w:rPr><w:t>第</w:t></w:r></w:p>')
        bod = P(f"{i+1}条　{SENT}")
        return cap + bod
    body = "".join(block(i) for i in range(30)) + pg.sectpr()
    pg.write_docx(out("probedropcap_dropcap.docx"), pg.doc(body))

# ---- 13. table inside a 2-column section ----------------------------------
def p_col_table():
    small = ('<w:tbl><w:tblPr><w:tblW w:w="3800" w:type="dxa"/>' + TBORD + '</w:tblPr>'
             '<w:tblGrid><w:gridCol w:w="3800"/></w:tblGrid>'
             + "".join(f'<w:tr><w:tc><w:tcPr><w:tcW w:w="3800" w:type="dxa"/></w:tcPr>{_cellp("項目" + str(j+1) + "：" + SENT[:24])}</w:tc></w:tr>' for j in range(6))
             + '</w:tbl>')
    body = conds(1, 20) + small + conds(21, 50) + pg.sectpr(cols='<w:cols w:num="2" w:space="425" w:sep="1"/>')
    pg.write_docx(out("probecoltbl_tableincols.docx"), pg.doc(body))

# ---- 14. very long single paragraph spanning 3 pages ----------------------
def p_long_para():
    big = P(SENT * 40)  # ~40x the sentence in ONE paragraph -> spans pages
    body = conds(1, 3) + big + conds(4, 6) + pg.sectpr()
    pg.write_docx(out("problongpara_longpara.docx"), pg.doc(body))

PROBES = [p_cols_mixed, p_cols_unequal, p_header_tall, p_footer_tall,
          p_table_header_repeat, p_table_cantsplit, p_table_vmerge,
          p_ruby_body, p_list_numbered, p_page_border, p_widow,
          p_keepnext, p_dropcap, p_col_table, p_long_para]

if __name__ == "__main__":
    for fn in PROBES:
        try:
            fn(); print("ok  ", fn.__name__)
        except Exception as e:
            print("FAIL", fn.__name__, e)
