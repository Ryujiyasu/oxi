"""S420: minimal repro to confirm Word's kinsoku REBALANCE for line-start-
prohibited chars (S409 hypothesis). Oxi force-fits a prohibited char onto
the overflow line; Word is thought to pull the preceding char down so the
prohibited char has a companion (ed025 pi=1374 "（× × ×）" → 5+2 split).

Builds a 3-cell-per-row table (tblCellMar=99) where each row's cell 0 has a
narrow tcW that forces a wrap, with text ending in a line-start-prohibited
char. COM-measures the per-character line assignment via the rendered Y
(chars sharing a Y are on the same visual line — Y is alignment-independent,
unlike Information(5) x; S416).

Test cases (cell tcW chosen so 7 fullwidth chars overflow, ~5 fit):
  K1  "（×　×　×）"  tcW=1500  — the ed025 pattern (paren + spaced ×)
  K2  "あいうえおか）" tcW=1500 — kana + closing paren
  K3  "あいうえおかき" tcW=1500 — NO prohibited char (natural wrap baseline)
  K4  "テスト文章です。" tcW=1500 — ending in 。(prohibited)
  K5  "（×　×　×）"  tcW=1900  — wider: does it still 2-line or fit 1?

Per case, report the line grouping (chars per rendered line) so we can see
whether Word's last line starts with the prohibited char alone (force-fit)
or with a companion (rebalance).

Instrumentation only.
"""
from __future__ import annotations
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

REPO = r'c:\Users\ryuji\oxi-main'
OUT = os.path.join(REPO, 'tools', 'golden-test', 'repros', 's420_kinsoku_rebalance.docx')

TESTS = [
    ('K1', '（×　×　×）', 1500),
    ('K2', 'あいうえ␣お）'.replace('␣', ''), 1500),
    ('K3', 'あいうえおかき', 1500),
    ('K4', 'テスト文章です。', 1500),
    ('K5', '（×　×　×）', 1900),
]


def row_xml(tcw, text):
    p = (f'<w:p><w:pPr><w:ind w:firstLineChars="100" w:firstLine="210"/>'
         f'<w:jc w:val="left"/></w:pPr>'
         f'<w:r><w:rPr><w:sz w:val="21"/></w:rPr>'
         f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>')
    tc0 = f'<w:tc><w:tcPr><w:tcW w:w="{tcw}" w:type="dxa"/></w:tcPr>{p}</w:tc>'
    filler = ('<w:tc><w:tcPr><w:tcW w:w="1600" w:type="dxa"/></w:tcPr>'
              '<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
              '<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>x</w:t></w:r></w:p></w:tc>')
    return f'<w:tr>{tc0}{filler}{filler}</w:tr>'


def build():
    rows = ''.join(row_xml(tcw, txt) for _, txt, tcw in TESTS)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="145" w:type="dxa"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           '<w:tblCellMar><w:left w:w="99" w:type="dxa"/><w:right w:w="99" w:type="dxa"/></w:tblCellMar>'
           '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="1900"/><w:gridCol w:w="1600"/><w:gridCol w:w="1600"/></w:tblGrid>'
           f'{rows}</w:tbl>')
    parts = {
     '[Content_Types].xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
     '_rels/.rels':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
     'word/_rels/document.xml.rels':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
     'word/styles.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults></w:styles>',
     'word/document.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+tbl+'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, c in parts.items():
            z.writestr(n, c)
    print(f'built {OUT}')


def measure():
    import win32com.client as wc
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(OUT, ReadOnly=True)
    doc.Repaginate()
    try:
        t = doc.Tables(1)
        for ri in range(1, t.Rows.Count + 1):
            cell = t.Cell(ri, 1)
            p = cell.Range.Paragraphs(1)
            rng = p.Range
            txt = (rng.Text or '').rstrip('\r\n\x07')
            s = rng.Start
            # per-char rendered Y (line grouping) — Y is alignment-independent
            lines = {}
            for i, ch in enumerate(txt):
                cr = doc.Range(s + i, s + i + 1)
                try:
                    y = round(float(cr.Information(6)), 1)
                except Exception:
                    y = -1.0
                lines.setdefault(y, []).append(ch)
            label = TESTS[ri - 1][0]
            ordered = [lines[y] for y in sorted(lines)]
            grouping = ' | '.join(''.join(g) for g in ordered)
            print(f'{label}: {len(ordered)} line(s): {grouping!r}')
    finally:
        doc.Close(SaveChanges=0)
        word.Quit()


if __name__ == '__main__':
    build()
    measure()
