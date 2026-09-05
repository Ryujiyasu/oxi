# -*- coding: utf-8 -*-
"""When does Word HANG a line-final 。 past the right margin (ぶら下げ)?

educational__08709ff2 (compat 11, compressPunctuation, noPunctuationKerning,
balanceSingleByteDoubleByteWidth, numbered paragraphs, BIZ UDPGothic 18pt,
line=576 exact): the row 「⑤検索結果の中から見たい項目をダブルタップします。」
has 「す」 ending 3.3pt before the margin and Word WRAPS 「す。」 instead of
hanging the 。 -- Oxi's S601 hangs it (the preceding content fits) and keeps
the line. Sweep the document's properties one at a time on a synthetic line
whose last text glyph ends just inside the margin, and read from Word's PDF
whether the 。 hangs (1 line) or wraps (2 lines).

    python _pb_hangpunct_gen.py gen
    python _pb_hangpunct_gen.py pdf      # Word truth
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_hangpunct")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

# content width: A4 11906 - 1701*2 = 8504tw = 425.2pt. 10.5pt: 40 chars = 420.0 (5.2 left),
# then 。 (10.5, compressed 5.25) -> hangs or wraps?  39 chars = 409.5 (15.7 left) -> 。 fits compressed? (5.25 -> 414.75 fits fully? 409.5+10.5=420 fits!) so use 40.
# (label, n_chars, compat, extra settings, numbered?, jc, face, sz half-points, csc)
ARMS = [
    ("base40_c15", 40, 15, "", False, "both", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("base40_c11", 40, 11, "", False, "both", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("left40_c15", 40, 15, "", False, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("left40_c11", 40, 11, "", False, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("nopk40_c11", 40, 11, "<w:noPunctuationKerning/>", False, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("bal40_c11", 40, 11, "<w:balanceSingleByteDoubleByteWidth/>", False, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("num40_c11", 40, 11, "", True, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("num40_c15", 40, 15, "", True, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("dnc40_c11", 40, 11, "", False, "left", "ＭＳ 明朝", 21, "doNotCompress"),
    ("dnc40_c15", 40, 15, "", False, "left", "ＭＳ 明朝", 21, "doNotCompress"),
    ("bizudp23_c11", 23, 11, "", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),   # 18pt: 23 chars = 414 (11.2 left)
    ("bizudp23_c15", 23, 15, "", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("msm23_c11_18", 23, 11, "", False, "left", "ＭＳ 明朝", 36, "compressPunctuation"),
    ("num38_c11", 38, 11, "", True, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("num38_c15", 38, 15, "", True, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("num38_both_c15", 38, 15, "", True, "both", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("exact40_c11", 40, 11, "EXACT", False, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("exact40_c15", 40, 15, "EXACT", False, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("bizsu22_c11", 22, 11, "SU", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("bizsu22_c15", 22, 15, "SU", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("bizTS21_c11", 21, 11, "TS", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),      # 21国+タッし = 424.45, 。 half 9 -> 518 > 510: hang?
    ("bizTS20_num_c11", 20, 11, "TS", True, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("bizTS20_num_all_c11", 20, 11, "TSALL", True, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("bizTS21_c15", 21, 15, "TS", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("dnwtwp40_c11", 40, 11, "<w:doNotWrapTextWithPunct/>", False, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("dnueabr40_c11", 40, 11, "<w:doNotUseEastAsianBreakRules/>", False, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("altkinsoku40_c11", 40, 11, "<w:useAltKinsokuLineBreakRules/>", False, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("dnwtwp40_both_c11", 40, 11, "<w:doNotWrapTextWithPunct/>", False, "both", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("dnwtwp40_both_c15", 40, 15, "<w:doNotWrapTextWithPunct/>", False, "both", "ＭＳ 明朝", 21, "compressPunctuation"),
    ("slice_c11", 0, 11, "SLICE", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("slice_nozw_c11", 0, 11, "SLICENOZW", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("slice_c15", 0, 15, "SLICE", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("slice_full_c11", 0, 11, "SLICEFULL", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("slice_halfA_c11", 0, 11, "SLICEA", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("slice_halfB_c11", 0, 11, "SLICEB", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("slice_ot_c11", 0, 11, "SLICEOT", False, "left", "BIZ UDPゴシック", 36, "compressPunctuation"),
    ("all_c11", 40, 11, "<w:noPunctuationKerning/><w:balanceSingleByteDoubleByteWidth/><w:useFELayout/><w:doNotLeaveBackslashAlone/>", True, "left", "ＭＳ 明朝", 21, "compressPunctuation"),
]


def docx(label):
    return os.path.join(OUT, "hangpunct_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')
    numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + NS + ">"
                 '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
                 '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimalEnclosedCircle"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/>'
                 '<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
                 '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
    for label, n, compat, extra, numbered, jc, face, sz, csc in ARMS:
        SL = ("SLICE", "SLICENOZW", "SLICEFULL", "SLICEA", "SLICEB", "SLICEOT")
        dsz = 36 if extra in SL else 21
        slice_mode = extra if extra in SL else None
        ea_face = "BIZ UDPゴシック" if slice_mode else "ＭＳ 明朝"
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="%s" w:hAnsi="Century"/>'
                    '<w:kern w:val="2"/><w:sz w:val="%d"/>' % (ea_face, dsz)
                  + '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
                    "<w:pPrDefault/></w:docDefaults>"
                    '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                    '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
        exact = extra in ("EXACT", "TSALL"); su = extra == "SU"; ts = extra in ("TS", "TSALL")
        if slice_mode:
            base_flags = "<w:noPunctuationKerning/><w:balanceSingleByteDoubleByteWidth/><w:useFELayout/><w:doNotLeaveBackslashAlone/><w:doNotWrapTextWithPunct/><w:doNotUseEastAsianBreakRules/><w:useAltKinsokuLineBreakRules/><w:doNotExpandShiftReturn/><w:doNotUseIndentAsNumberingTabStop/>"
            halfA = "<w:spaceForUL/><w:ulTrailSpace/><w:adjustLineHeightInTable/><w:useNormalStyleForList/><w:allowSpaceOfSameStyleInTable/><w:doNotSuppressIndentation/><w:doNotAutofitConstrainedTables/><w:autofitToFirstFixedWidthCell/>"
            halfB = "<w:displayHangulFixedWidth/><w:splitPgBreakAndParaMark/><w:doNotVertAlignCellWithSp/><w:doNotBreakConstrainedForcedTable/><w:doNotVertAlignInTxbx/><w:useAnsiKerningPairs/><w:cachedColBalance/>"
            ot = '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>'
            extra = base_flags + {"SLICEFULL": halfA + halfB, "SLICEA": halfA, "SLICEB": halfB, "SLICEOT": halfA + halfB + ot}.get(slice_mode, "")
        if extra == "TSALL":
            extra = "<w:noPunctuationKerning/><w:balanceSingleByteDoubleByteWidth/><w:useFELayout/><w:doNotLeaveBackslashAlone/>"
        elif extra in ("EXACT", "SU", "TS"):
            extra = ""
        tabstop = '<w:defaultTabStop w:val="420"/>' if slice_mode else ""
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                    + '<w:characterSpacingControl w:val="%s"/>%s<w:compat>%s'
                      '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/>'
                      "</w:compat></w:settings>" % (csc, tabstop, extra, compat))
        rpr = '<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr>' % (face, face, face, sz)
        numpr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if numbered else ""
        # the numbered arm loses 18pt (the hanging indent) -> one char fewer
        nn = n
        body = ""
        if slice_mode:
            import re as _re
            dz = zipfile.ZipFile(os.path.join(REPO, "pipeline_data", "docx_corpus", "ja", "educational", "08709ff2e7c5fdfa.docx"))
            dd = dz.read("word/document.xml").decode("utf-8")
            i = dd.find("見たい項目をダブルタップ"); a = dd.rfind("<w:p ", 0, i); b = dd.find("</w:p>", i) + 6
            para = dd[a:b].replace('w:numId w:val="2"', 'w:numId w:val="1"')
            para = _re.sub(r' (?:w14|w16\w*|w15):\w+="[^"]*"', '', para)
            para = _re.sub(r' w:rsid\w*="[^"]*"', '', para)
            if slice_mode == "SLICENOZW":
                para = _re.sub(r"<w:r>(?:(?!</w:r>).)*?<w:t>​</w:t></w:r>", "", para)
            body += para * 3 + "<w:p><w:r><w:t>末尾</w:t></w:r></w:p>"
        for k in range(0 if slice_mode else 3):
            text = "国" * nn + ("す" if su else "") + ("タッし" if ts else "") + "。"
            sp = '<w:spacing w:line="576" w:lineRule="exact"/>' if exact else ""
            body += ('<w:p><w:pPr>%s%s<w:jc w:val="%s"/>%s</w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>' % (numpr, sp, jc, rpr, rpr, text))
        body += "<w:p><w:r><w:t>末尾</w:t></w:r></w:p>"
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/>'
                 '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                       '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                       '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
                       "</Relationships>")
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/numbering.xml", numbering)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def pdf():
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for label, *_ in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                d.SaveAs2(docx(label)[:-5] + ".word.pdf", 17)
            finally:
                d.Close(False)
    finally:
        app.Quit()
    print("== WORD (PDF): rows per paragraph (1 = the 。 hangs / fits, 2 = wraps); x of the 。 and the last 国 end; right margin = 510.2 ==")
    for label, n, compat, extra, numbered, jc, face, sz, csc in ARMS:
        pg = fitz.open(docx(label)[:-5] + ".word.pdf")[0]
        rows = []
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                chars = [c for sp in l["spans"] for c in sp["chars"] if c["c"].strip()]
                if chars:
                    rows.append((round(chars[0]["origin"][1], 1), "".join(c["c"] for c in chars), [(c["c"], round(c["origin"][0], 1), round(c["bbox"][2], 1)) for c in chars if c["c"] == "。"], round(chars[-1]["bbox"][2], 1)))
        rows.sort()
        body = [r for r in rows if "国" in r[1] or r[1] == "。"]
        n_rows_first = 0
        for r in rows:
            if "国" in r[1] or r[1] == "。":
                n_rows_first += 1
            if n_rows_first and r[1] == "。":
                break
        first_para_rows = rows[:2]
        print("%-16s n=%d c%d num=%-5s jc=%-4s %-8s %s -> rows=%d | row0 end=%.1f maru=%s | row1=%r" % (
            label, n, compat, numbered, jc, face, csc[:6], len(body), rows[0][3], rows[0][2], rows[1][1][:6] if len(rows) > 1 else None))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    else:
        gen()
