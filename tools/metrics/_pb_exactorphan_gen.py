# -*- coding: utf-8 -*-
"""A 2-line paragraph with EXACT line spacing at the page bottom: does Word
keep its first line when only the NATURAL height of the second would fit?

technical__898a80 p11 (BIZ UDPGothic 18pt, line=576 exact, widowControl on):
the 2-line paragraph 「②画面上では…」 starts at 704.3 with the content bottom at
756.85 -- 704.3 + 28.8 + 28.8 = 761.9 does not fit, 704.3 + 28.8 + natural
23.4 = 756.5 does. Word (today's PDF) pushes the whole paragraph; Oxi's S608
look-ahead (natural height for the last line, derived on x2.0 multiples) keeps
line 0. Sweep the paragraph's start against the bottom and read, through COM,
which page each of its two lines lands on.

    python _pb_exactorphan_gen.py gen
    python _pb_exactorphan_gen.py com
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_exactorphan")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

# A4, top 1985 (99.25) bottom 1701 (85.05): band 99.25 .. 756.9 = 657.6pt.
# exact 28.8 lines: 22 fillers = 633.6 -> the 2-line para starts at 732.85 (24 left: neither line fits -> push)
# The interesting band: the para starts where 1 exact + natural fits but 2 exact do not:
#   start s: s + 28.8 + nat(23.4) <= 756.9  <=> s <= 704.7 ;  s + 57.6 > 756.9 <=> s > 699.3
# Use a spacer paragraph of exact height H before the 2-liner to place s finely.
# (label, n_filler_lines, spacer exact height in twips or 0, face, sz half-points, exact twips)
ARMS = [
    ("s699", 20, 0, "BIZ UDPゴシック", 36, 576),      # 20*28.8 = 576 -> s = 675.25 (+spacer)
    ("s700", 20, 500, "BIZ UDPゴシック", 36, 576),    # spacer 25.0 -> s = 700.25: 2 exact = 757.85 > 756.9; 1 exact + nat = 752.5 fits
    ("s702", 20, 540, "BIZ UDPゴシック", 36, 576),    # spacer 27.0 -> s = 702.25
    ("s704", 20, 580, "BIZ UDPゴシック", 36, 576),    # spacer 29.0 -> s = 704.25 (the document's 704.3)
    ("s706", 20, 620, "BIZ UDPゴシック", 36, 576),    # spacer 31.0 -> s = 706.25: 1 exact + nat = 758.4 > bottom -> push either way
    ("s698", 20, 460, "BIZ UDPゴシック", 36, 576),    # spacer 23.0 -> s = 698.25: 2 exact = 755.85 fits -> keep both
    ("s702_msg", 20, 540, "ＭＳ ゴシック", 36, 576),
    ("s702_msm105", 20, 540, "ＭＳ 明朝", 21, 576),   # 10.5pt text in a 28.8 exact line: natural 13.6
]
TEXT2 = "②画面上ではシリが応答状態のままとなっていますので、ホーム画面に戻る操作を行いシリを終了します。"


def docx(label):
    return os.path.join(OUT, "exactorphan_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    for label, nfill, spacer, face, sz, exact in ARMS:
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="%s" w:hAnsi="Century"/>'
                  '<w:kern w:val="2"/><w:sz w:val="%d"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
                  "<w:pPrDefault/></w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:widowControl/><w:jc w:val="left"/><w:spacing w:line="%d" w:lineRule="exact"/></w:pPr></w:style></w:styles>'
                  % (face, sz, exact))
        body = "".join("<w:p><w:r><w:t>行%d</w:t></w:r></w:p>" % i for i in range(nfill))
        if spacer:
            body += '<w:p><w:pPr><w:spacing w:line="%d" w:lineRule="exact"/></w:pPr><w:r><w:t>間</w:t></w:r></w:p>' % spacer
        body += "<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % TEXT2
        body += "<w:p><w:r><w:t>次の段落</w:t></w:r></w:p>"
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/>'
                 '<w:docGrid w:type="lines" w:linePitch="576"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
            z.writestr("word/styles.xml", styles)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def com():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    print("== WORD (COM): the 2-line paragraph's start y (Info6), page of its first and last char, page of the next paragraph ==")
    try:
        for label, nfill, spacer, face, sz, exact in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                n = d.Paragraphs.Count
                para = d.Paragraphs(n - 1)
                r = para.Range
                s = d.Range(r.Start, r.Start)
                e = d.Range(r.End - 2, r.End - 2)
                nxt = d.Paragraphs(n).Range
                ns = d.Range(nxt.Start, nxt.Start)
                print("%-12s %-8s %4.1fpt exact=%.1f -> start y=%.2f p%d, last char p%d (y=%.2f), next para p%d" % (
                    label, face, sz / 2, exact / 20, s.Information(6), s.Information(3), e.Information(3), e.Information(6), ns.Information(3)))
            finally:
                d.Close(False)
    finally:
        app.Quit()


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "com":
        com()
    else:
        gen()
