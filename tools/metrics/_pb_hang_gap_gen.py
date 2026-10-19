# -*- coding: utf-8 -*-
"""Does Word hang a line-end mark that would START before the right edge?

ja/educational__08709ff2 «…ダブルタップします。» (BIZ UDPGothic 18pt, jc=left): す
ends 3.4pt short of the edge, 。 (10.8) does not fit -- Word does NOT hang it, it
pushes «す。» to the next line.  The same document's «…いただけますと、» hangs its 、
only with the 、 starting within 0.4pt of the edge; sweeping the right indent 0.2pt
at a time, Word squeezes the line to keep the 、 there for 6.2pt, then pushes «と、».

Arms: face x jc, one page per right indent (0.4pt steps over one em), text = kana
then a mark then more kana.  Per page: the first line's last two characters, the
mark's start minus the edge, and whether it hangs.

    python tools/metrics/_pb_hang_gap_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_hang_gap")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
FACES = os.environ.get("PB_FACES", "BIZ UDPGothic,MS Gothic").split(",")
JCS = os.environ.get("PB_JCS", "left,both").split(",")
MARK = os.environ.get("PB_MARK", "。")
N = int(os.environ.get("PB_N", "22"))
# PB_TAIL=0: the mark ends the paragraph (08709ff2 «…します。¶»)
TAIL = "" if os.environ.get("PB_TAIL") == "0" else "まみむめもやゆよ"
# PB_OPEN=k: an opening 「 before the k-th last kana (08709ff2 «…押し「バツ印」を»)
KANA = os.environ.get("PB_KANA", "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほ")[:N]
OPEN = int(os.environ.get("PB_OPEN", "0"))
TEXT = (KANA[:-OPEN] + "「" + KANA[-OPEN:] if OPEN else KANA) + MARK + TAIL
RIGHTS = list(range(0, 400, 8))
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
CM = int(os.environ.get("PB_CM", "11"))
# PB_RPR: extra run properties (additive bisection against 08709ff2, e.g. <w:kern w:val="2"/>)
RPR_EXTRA = os.environ.get("PB_RPR", "")
PPR_EXTRA = os.environ.get("PB_PPR", "")
# PB_LANGVAL=en-US: the run's Latin language (08709ff2's docDefaults say en-US)
LANGVAL = os.environ.get("PB_LANGVAL", "")
# PB_COMPAT="<w:doNotWrapTextWithPunct/>": extra w:compat children (08709ff2's legacy flags)
COMPAT_EXTRA = os.environ.get("PB_COMPAT", "")
SECT_EXTRA = os.environ.get("PB_SECT", "")
CSC = os.environ.get("PB_CSC", "compressPunctuation")
LEFT_MARGIN, PAGE_W, RIGHT_MARGIN = 1701, 11906, 1701


def build(face, jc):
    ps = []
    for k, r in enumerate(RIGHTS):
        ps.append(('<w:p><w:pPr>%s' + PPR_EXTRA.replace("%", "%%") + '<w:ind w:right="%d"/><w:jc w:val="%s"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
                   '<w:sz w:val="36"/>' + RPR_EXTRA.replace("%", "%%") + '<w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>%s</w:t></w:r></w:p>')
                  % ("<w:pageBreakBefore/>" if k else "", r, jc, face, face, face, TEXT))
    sect = ('<w:sectPr><w:pgSz w:w="%d" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="%d" w:right="%d"/>'
            % (PAGE_W, LEFT_MARGIN, RIGHT_MARGIN)) + SECT_EXTRA + '</w:sectPr>'
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s%s</w:body></w:document>'
           % (W, "".join(ps), sect))
    st = (('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:characterSpacingControl w:val="%s"/><w:compat>'
           + COMPAT_EXTRA.replace("%", "%%") +
           '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>') % (W, CSC, CM))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, ("x%d_" % (abs(hash(RPR_EXTRA + PPR_EXTRA + SECT_EXTRA + LANGVAL + COMPAT_EXTRA)) % 10000) if RPR_EXTRA + PPR_EXTRA + SECT_EXTRA + LANGVAL + COMPAT_EXTRA else "") + "hg_%s_%s_%s_c%d_%s%s.docx" % (face.replace(" ", ""), jc, "%x" % ord(MARK), CM, CSC[:4], "" if TAIL else "_end") if not OPEN else "hg_%s_%s_%x_c%d_open%d.docx" % (face.replace(" ", ""), jc, ord(MARK), CM, OPEN))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/settings.xml", st)
        z.writestr("word/document.xml", doc)
    return path


def report(pdf):
    import fitz
    rows = []
    for r, pg in zip(RIGHTS, fitz.open(pdf)):
        edge = (PAGE_W - RIGHT_MARGIN - r) / 20
        lines = {}
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        if c["c"].strip():
                            lines.setdefault(round(c["origin"][1]), []).append(c)
        first = sorted(lines[min(lines)], key=lambda c: c["origin"][0])
        last = first[-1]
        rows.append((r, edge, "".join(c["c"] for c in first[-2:]), round(last["bbox"][0] - edge, 2), last["bbox"][2] > edge + 0.05))
    return rows


if __name__ == "__main__":
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for face in FACES:
            for jc in JCS:
                docx = build(face, jc); pdf = docx[:-5] + ".pdf"
                d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
                rows = report(pdf)
                print("== %s jc=%s mark=%s c%d %s%s" % (face, jc, MARK, CM, CSC, "" if TAIL else " PARA-END"), flush=True)
                prev = None
                for r, edge, tail, gap, hung in rows:
                    key = (tail, hung)
                    if key != prev:
                        print("   right=%4d edge=%.2f  tail=%s  last.start-edge=%+.2f  %s" % (r, edge, tail, gap, "HANG" if hung else ""), flush=True)
                        prev = key
    finally:
        w.Quit()
