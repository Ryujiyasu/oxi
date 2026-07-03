# -*- coding: utf-8 -*-
"""Adversarial structural probe batch 5 (FINAL sweep) — the remaining
candidates from the batch-4 coverage map + newly enumerated untested classes.

  - tracked changes w:ins/w:del (ground truth = whatever Word's default
    markup view paginates; del segments placed AFTER the 30-char prefix
    window so matching stays clean)
  - TABLE STYLE formatting (tblStyle cellMar/borders from styles.xml, no
    direct formatting — modern-doc pattern)
  - w:object legacy embed (VML presentation shape; no OLE binary — tests
    the w:object wrapper parse path without repair risk)
  - modern inline wps textbox via mc:AlternateContent (Choice=wps drawing,
    Fallback=VML) — THE modern textbox form
  - footnote pos=beneathText; footnote refs INSIDE table cells
  - RTL Arabic body (w:bidi + w:rtl)
  - hidden runs (w:vanish) mid-paragraph — Word hides them from layout
  - autoHyphenation with Latin tokens embedded in CJK paragraphs
  - extreme font sizes (6pt / 28pt) in a typed grid
  - two ADJACENT tables crossing a page break
  - page-anchored framePr frame (vAnchor/hAnchor=page)
  - inline IMAGE in the header (the estimate's image arm)
  - wrap-control characters: NBSP / ZWSP / soft hyphen / w:noBreakHyphen
  - w:hideMark thin spacer rows (trHeight exact + hidden cell mark)

Deliberately EXCLUDED (justified): vAlign (pagination-neutral),
mirrorMargins±gutter (content width invariant), line numbering (margin-only),
comments (balloons don't change print pagination), PAGE/NUMPAGES field width
(cache-driven on both sides per S708), numbering restart (marker-width-only,
blocked on the problist fix), real OLE binaries (repair-prompt risk),
lrV/rlV section text direction (rare).

All doc_ids start with "probeq" (unique prefix; "probew" would collide with
batch-2 probewidow).

Run: python tools/metrics/_probe_gen5.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import _probe_gen as pg
import _probe_gen3 as g3
import _probe_gen4 as g4

MINCHO = pg.MINCHO
esc = pg.esc
SENT = pg.SENT
rpr, P, cond, conds, brk_para = g3.rpr, g3.P, g3.cond, g3.conds, g3.brk_para

def out(n): return pg.out(n)

HDR_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
HDR_RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
FN_CT = ('<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',)
FN_REL = ('<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',)

def footnotes_xml(n, label="注記"):
    fnrpr = rpr("18")
    def note(i):
        return (f'<w:footnote w:id="{i+2}"><w:p><w:pPr><w:rPr>{fnrpr}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{fnrpr}<w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
                f'<w:r><w:rPr>{fnrpr}</w:rPr><w:t xml:space="preserve">{label}{i+1}：本条の適用に関する補足説明であって、実務上の取扱いを示すものである。</w:t></w:r></w:p></w:footnote>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:footnotes {pg.DOC_NS}>'
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
            + "".join(note(i) for i in range(n)) + '</w:footnotes>')

# ---- 1. tracked changes: w:ins + w:del runs ----------------------------------
def p_insdel():
    r = rpr()
    parts = []
    for i in range(1, 41):
        if i % 2 == 0:
            # head (>=35 chars incl 第N条 prefix) keeps the match window clean
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT[:35])}</w:t></w:r>'
                f'<w:ins w:id="{100+i}" w:author="probe" w:date="2026-07-01T00:00:00Z">'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">（挿入：追加された規定の細目をここに定める）</w:t></w:r></w:ins>'
                f'<w:del w:id="{500+i}" w:author="probe" w:date="2026-07-01T00:00:00Z">'
                f'<w:r><w:rPr>{r}</w:rPr><w:delText xml:space="preserve">（削除：旧規定の細目であった部分）</w:delText></w:r></w:del>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(SENT[35:])}</w:t></w:r></w:p>')
        else:
            parts.append(cond(i))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probeqinsdel_trackedchanges.docx"), pg.doc(body))

# ---- 2. TABLE STYLE cellMar/borders (no direct formatting) --------------------
def p_tblstyle():
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
              '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
              '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>'
              '<w:style w:type="table" w:styleId="PT"><w:name w:val="ProbeTable"/>'
              '<w:tblPr>'
              '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="240" w:type="dxa"/>'
              '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="240" w:type="dxa"/></w:tblCellMar>'
              '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
              '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
              '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
              '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
              '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
              '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
              '</w:tblPr></w:style></w:styles>')
    def cellp(txt):
        r = rpr()
        return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')
    rows = "".join(
        '<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="2200" w:type="dxa"/></w:tcPr>{cellp(f"第{i+1}項")}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="6800" w:type="dxa"/></w:tcPr>{cellp(SENT)}</w:tc>'
        '</w:tr>' for i in range(40))
    tbl = ('<w:tbl><w:tblPr><w:tblStyle w:val="PT"/><w:tblW w:w="9000" w:type="dxa"/>'
           '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" '
           'w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="2200"/><w:gridCol w:w="6800"/></w:tblGrid>'
           + rows + '</w:tbl>')
    body = conds(1, 3) + tbl + conds(4, 6) + pg.sectpr()
    pg.write_docx(out("probeqtblstyle_tablestyle.docx"), pg.doc(body),
                  extra_parts={"word/styles.xml": styles})

# ---- 3. w:object legacy embed (VML presentation shape) ------------------------
def p_object():
    r = rpr()
    def obj_para(i):
        return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr>'
                '<w:object w:dxaOrig="2880" w:dyaOrig="1440">'
                '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" '
                f'id="OBJ{i}" style="width:144pt;height:72pt">'
                '<v:imagedata r:id="rId20"/></v:shape></w:object></w:r></w:p>')
    parts = []
    for i in range(1, 41):
        parts.append(cond(i))
        if i % 10 == 0:
            parts.append(obj_para(i))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probeqobject_oleobjectshape.docx"), pg.doc(body),
                  extra_parts=dict(g3.PNG_PART), ct_extra=g3.PNG_CT, rel_extra=g3.PNG_REL)

# ---- 4. modern inline wps textbox via mc:AlternateContent ----------------------
def p_wps_txbx():
    r = rpr()
    inner = P("箱内注記：見本", jc="left", sz="18")
    def wps_inline(pid):
        cx, cy = 1780000, 640000  # ~140pt x 50pt
        choice = (
            f'<w:drawing><wp:inline {g3.WP_NS} distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{pid}" name="TB{pid}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic {g3.A_NS}>'
            '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp>'
            '<wps:cNvSpPr/><wps:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="F0F0F0"/></a:solidFill>'
            '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
            '</wps:spPr>'
            f'<wps:txbx><w:txbxContent>{inner}</w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" '
            'rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
            '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>')
        fallback = ('<w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" '
                    'style="width:140pt;height:50pt" fillcolor="#f0f0f0" strokecolor="black">'
                    f'<v:textbox><w:txbxContent>{inner}</w:txbxContent></v:textbox></v:rect></w:pict>')
        # xmlns:wps must be in scope AT the AlternateContent element so the
        # mc:Choice Requires="wps" prefix resolves (Word rejects it otherwise)
        return ('<mc:AlternateContent '
                'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
                f'<mc:Choice Requires="wps">{choice}</mc:Choice>'
                f'<mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>')
    parts = []
    for i in range(1, 41):
        if i % 9 == 0:
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT[:30])}</w:t></w:r>'
                f'<w:r>{wps_inline(i)}</w:r>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(SENT[30:])}</w:t></w:r></w:p>')
        else:
            parts.append(cond(i))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probeqwps_modernwpstextbox.docx"), pg.doc(body))

# ---- 5. footnote pos=beneathText -----------------------------------------------
def p_fn_beneath():
    n = 22
    r = rpr()
    sup = rpr(extra='<w:vertAlign w:val="superscript"/>')
    def fn_para(i):
        return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i+1}条　{esc(SENT)}</w:t></w:r>'
                f'<w:r><w:rPr>{sup}</w:rPr><w:footnoteReference w:id="{i+2}"/></w:r></w:p>')
    body = ("".join(fn_para(i) for i in range(n))
            + pg.sectpr(sect_type='<w:footnotePr><w:pos w:val="beneathText"/></w:footnotePr>'))
    pg.write_docx(out("probeqfnbeneath_footnotebeneathtext.docx"), pg.doc(body),
                  extra_parts={"word/footnotes.xml": footnotes_xml(n)},
                  ct_extra=FN_CT, rel_extra=FN_REL)

# ---- 6. footnote refs INSIDE table cells ----------------------------------------
def p_fn_in_cell():
    n = 20
    r = rpr()
    sup = rpr(extra='<w:vertAlign w:val="superscript"/>')
    def cellp(i):
        return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i+1}項：{esc(SENT[:65])}</w:t></w:r>'
                f'<w:r><w:rPr>{sup}</w:rPr><w:footnoteReference w:id="{i+2}"/></w:r></w:p>')
    rows = "".join(
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{cellp(i)}</w:tc></w:tr>'
        for i in range(n))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + g3.pg2_borders() + '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' + rows + '</w:tbl>')
    body = conds(1, 3) + tbl + conds(4, 6) + pg.sectpr()
    pg.write_docx(out("probeqfncell_footnoteincell.docx"), pg.doc(body),
                  extra_parts={"word/footnotes.xml": footnotes_xml(n, "表注")},
                  ct_extra=FN_CT, rel_extra=FN_REL)

# ---- 7. RTL Arabic body (w:bidi + w:rtl) ----------------------------------------
AR_SENT = ("تُطبق أحكام هذه المادة بحسن نية وبما يتفق مع الأنظمة واللوائح المعمول بها، "
           "ويُعالج أي أمر ينشأ عن تنفيذها فوراً وبإنصاف ودون تأخير غير مبرر من قبل الأطراف المعنية.")

def p_bidi():
    def ap(i):
        rp = ('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
              '<w:sz w:val="22"/><w:szCs w:val="22"/><w:rtl/>')
        return (f'<w:p><w:pPr><w:bidi/><w:rPr>{rp}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{rp}</w:rPr><w:t xml:space="preserve">البند {i} — {AR_SENT}</w:t></w:r></w:p>')
    body = "".join(ap(i) for i in range(1, 46)) + pg.sectpr(grid='')
    pg.write_docx(out("probeqbidi_arabicrtl.docx"), pg.doc(body),
                  font="Arial", sz="22", cpunct=False)

# ---- 8. hidden runs (w:vanish) mid-paragraph ------------------------------------
def p_vanish_runs():
    r = rpr()
    hid = rpr(extra='<w:vanish/>')
    parts = []
    for i in range(1, 46):
        if i % 2 == 0:
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT[:35])}</w:t></w:r>'
                f'<w:r><w:rPr>{hid}</w:rPr><w:t xml:space="preserve">（隠し文字：この部分は表示も印刷もされない補足である）</w:t></w:r>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(SENT[35:])}</w:t></w:r></w:p>')
        else:
            parts.append(cond(i))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probeqvanish_hiddenruns.docx"), pg.doc(body))

# ---- 9. autoHyphenation + Latin tokens inside CJK paragraphs --------------------
def p_hyph_jp():
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:autoHyphenation/><w:hyphenationZone w:val="360"/>'
                '<w:characterSpacingControl w:val="compressPunctuation"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    def hp(i):
        return P(f"第{i}条　{g4.LONGWORDS}の要件を満たす場合に限り、{SENT[:55]}")
    body = "".join(hp(i) for i in range(1, 51)) + pg.sectpr()
    pg.write_docx(out("probeqhyphjp_hyphenationjp.docx"), pg.doc(body),
                  extra_parts={"word/settings.xml": settings})

# ---- 10. extreme font sizes (6pt / 28pt) in a typed grid ------------------------
def p_sizes():
    parts = []
    for i in range(1, 51):
        if i % 2 == 0:
            parts.append(P(f"第{i}条　{SENT[:70]}", sz="12"))   # 6pt
        else:
            parts.append(P(f"第{i}条　{SENT[:40]}", sz="56"))   # 28pt
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probeqsizes_extremesizes.docx"), pg.doc(body))

# ---- 11. two ADJACENT tables crossing a page break ------------------------------
def p_tbl_tbl():
    def cellp(txt):
        r = rpr()
        return (f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')
    def mktbl(tag, n):
        rows = "".join(
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{cellp(f"{tag}{j+1}：" + SENT[:55])}</w:tc></w:tr>'
            for j in range(n))
        return ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + g3.pg2_borders() + '</w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' + rows + '</w:tbl>')
    # two adjacent tables (no paragraph between), then a third small one
    body = (conds(1, 3) + mktbl("甲第", 22) + mktbl("乙第", 22)
            + conds(4, 5) + mktbl("丙第", 5) + conds(6, 7) + pg.sectpr())
    pg.write_docx(out("probeqtbltbl_adjacenttables.docx"), pg.doc(body))

# ---- 12. page-anchored framePr frame --------------------------------------------
def p_frame_page():
    r = rpr()
    fp = ('<w:framePr w:w="2835" w:h="2268" w:hRule="atLeast" w:hSpace="141" '
          'w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="6800" w:y="4500"/>')
    frame = (f'<w:p><w:pPr>{fp}<w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
             f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">（頁固定枠）別紙様式を参照のこと。</w:t></w:r></w:p>')
    body = conds(1, 2) + frame + conds(3, 48) + pg.sectpr()
    pg.write_docx(out("probeqframepg_pageanchoredframe.docx"), pg.doc(body))

# ---- 13. inline IMAGE in the header ---------------------------------------------
# The image rel (rId20) must live in header1.xml.rels, not the document rels.
def p_hdr_image_fixed():
    hdr_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
                '</Relationships>')
    hdr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:hdr {pg.DOC_NS}>'
           f'<w:p><w:r>{g3.inline_img(300, 1440000, 1080000, "rId20")}</w:r></w:p>'
           + P("（社章の下のヘッダ行）", jc="left") + '</w:hdr>')
    ref = '<w:headerReference w:type="default" r:id="rId10"/>'
    body = conds(1, 55) + pg.sectpr(sect_type=ref)
    pg.write_docx(out("probeqhdrimg_imageinheader.docx"), pg.doc(body),
                  extra_parts={"word/header1.xml": hdr,
                               "word/_rels/header1.xml.rels": hdr_rels,
                               **g3.PNG_PART},
                  ct_extra=(f'<Override PartName="/word/header1.xml" ContentType="{HDR_CT}"/>',) + g3.PNG_CT,
                  rel_extra=(f'<Relationship Id="rId10" Type="{HDR_RT}" Target="header1.xml"/>',))

# ---- 14. wrap-control characters: NBSP / ZWSP / SHY / noBreakHyphen -------------
def p_break_chars():
    r = rpr()
    parts = []
    for i in range(1, 49):
        k = i % 4
        if k == 0:
            # w:noBreakHyphen element between Latin halves
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　check</w:t></w:r>'
                '<w:r><w:noBreakHyphen/></w:r>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">list（分離不可ハイフン）に基づき、{esc(SENT[:55])}</w:t></w:r></w:p>')
        elif k == 1:
            txt = (f"第{i}条　rate limit value（NBSP連結）"
                   f"を超えない範囲で、{SENT[:55]}")
            parts.append(P(txt))
        elif k == 2:
            txt = (f"第{i}条　inter​governmental​responsibility"
                   f"（ZWSP挿入）につき、{SENT[:55]}")
            parts.append(P(txt))
        else:
            txt = (f"第{i}条　imple­menta­tion"
                   f"（ソフトハイフン）の手続は、{SENT[:55]}")
            parts.append(P(txt))
    body = "".join(parts) + pg.sectpr()
    pg.write_docx(out("probeqbrkchars_wrapcontrolchars.docx"), pg.doc(body))

# ---- 15. w:hideMark thin spacer rows --------------------------------------------
def p_hidemark():
    r = rpr()
    def crow(i):
        return ('<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i+1}項：{esc(SENT[:55])}</w:t></w:r></w:p></w:tc></w:tr>')
    thin = ('<w:tr><w:trPr><w:trHeight w:val="100" w:hRule="exact"/></w:trPr>'
            '<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/><w:hideMark/></w:tcPr>'
            f'<w:p><w:pPr><w:rPr>{rpr("2")}</w:rPr></w:pPr></w:p></w:tc></w:tr>')
    rows = "".join(crow(i) + thin for i in range(30))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + g3.pg2_borders() + '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' + rows + '</w:tbl>')
    body = conds(1, 2) + tbl + conds(3, 4) + pg.sectpr()
    pg.write_docx(out("probeqhidemark_thinspacerrows.docx"), pg.doc(body))

# ---- 15b. hideMark WITHOUT trHeight (isolates the hideMark semantics —
# v1 conflated it with hRule=exact, which clamps regardless; a hideMark cell
# with an auto row collapses to ~borders-only in Word) -----------------------
def p_hidemark2():
    r = rpr()
    def crow(i):
        return ('<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="left"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i+1}項：{esc(SENT[:55])}</w:t></w:r></w:p></w:tc></w:tr>')
    thin = ('<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/><w:hideMark/></w:tcPr>'
            f'<w:p><w:pPr><w:rPr>{rpr("2")}</w:rPr></w:pPr></w:p></w:tc></w:tr>')
    rows = "".join(crow(i) + thin for i in range(28))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>' + g3.pg2_borders() + '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' + rows + '</w:tbl>')
    body = conds(1, 2) + tbl + conds(3, 4) + pg.sectpr()
    pg.write_docx(out("probeqhidemk2_hidemarkauto.docx"), pg.doc(body))

PROBES = [p_insdel, p_tblstyle, p_object, p_wps_txbx, p_fn_beneath,
          p_fn_in_cell, p_bidi, p_vanish_runs, p_hyph_jp, p_sizes,
          p_tbl_tbl, p_frame_page, p_hdr_image_fixed, p_break_chars,
          p_hidemark, p_hidemark2]

if __name__ == "__main__":
    for fn in PROBES:
        try:
            fn(); print("ok  ", fn.__name__)
        except Exception as e:
            print("FAIL", fn.__name__, e)
