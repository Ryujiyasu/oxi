"""S440: faithful in-cell wrap-boundary measurement infrastructure.

Motivation (S434-S439): d77a item J wraps to 2 lines in Word but 1 in Oxi,
and FIVE hypotheses (char-width, szCs, kinsoku, margins, cell_w-vs-inner_w)
were all cleanly rejected. The residual is a ~20pt cell-wrap BUDGET gap that
could not be pinned because (a) hand-built repros were UNFAITHFUL (the
hanging indent did not apply) and (b) Word's true in-cell wrap boundary was
never measured directly.

This tool provides the missing infrastructure:

1. extract_cell_table(docx, match_text) -> a FAITHFUL standalone docx
   containing ONLY the table that encloses the matched paragraph, copied
   VERBATIM (tcPr/tblPr/tblGrid/pPr/rPr all preserved), with the document's
   styles.xml and sectPr, and the FULL Word namespace set declared (the
   missing w14:/w15:/mc: declarations are why earlier extractions failed to
   open in Word). This reproduces Word's exact rendering of the cell.

2. Faithfulness check: render with Oxi and (optionally) COM-measure Word;
   confirm the target paragraph's line count matches the real doc before
   trusting any probe.

3. Probe mode (--probe N): replace the target paragraph's body with 「本」×N
   (known 10.5pt full-width advance) keeping the real pPr (indent/marker
   path). Sweeping N and finding each engine's wrap flip-point yields the
   EXACT wrap budget for that cell, with zero mixed-width / pixel-rounding
   noise.

Run:
  python tools/metrics/cell_wrap_probe.py extract <doc_id> "<match text>"
  python tools/metrics/cell_wrap_probe.py extract d77a ウェブサイト全体
"""
from __future__ import annotations
import os, re, sys, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
OUT = os.path.join(REPO, "tools", "golden-test", "repros", "d77a_cellwrap")

# Full Word namespace set — omitting w14/w15/mc is why hand extractions failed
# to open ("use text recovery converter"). Declare them all on the root.
NS = ' '.join([
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
    'xmlns:o="urn:schemas-microsoft-com:office:office"',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
    'xmlns:v="urn:schemas-microsoft-com:vml"',
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"',
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
    'xmlns:w10="urn:schemas-microsoft-com:office:word"',
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"',
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"',
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"',
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    'mc:Ignorable="w14 w15 wp14"',
])
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')


def doc_id_to_path(doc_id):
    for f in os.listdir(DOCS):
        if f.lower().endswith(".docx") and f.startswith(doc_id):
            return os.path.join(DOCS, f)
    raise SystemExit(f"no docx for {doc_id}")


def enclosing_table(x, pos):
    """Smallest <w:tbl>..</w:tbl> span enclosing char offset pos."""
    ev = sorted([(m.start(), 'o') for m in re.finditer(r'<w:tbl>', x)] +
                [(m.end(), 'c') for m in re.finditer(r'</w:tbl>', x)])
    stack, best = [], None
    for p, k in ev:
        if k == 'o':
            stack.append(p)
        else:
            op = stack.pop()
            if op < pos < p and (best is None or op > best[0]):
                best = (op, p)
    return best


def extract(doc_id, match_text, probe=None):
    z = zipfile.ZipFile(doc_id_to_path(doc_id))
    x = z.read('word/document.xml').decode('utf-8')
    styles = z.read('word/styles.xml').decode('utf-8')
    # locate match: a position whose stripped context contains match_text
    pos = None
    needle = match_text[:4]
    for m in re.finditer(re.escape(needle), x):
        seg = re.sub(r'<[^>]+>', '', x[max(0, m.start()-300):m.start()+300])
        if match_text in seg:
            pos = m.start(); break
    if pos is None:
        raise SystemExit(f"match {match_text!r} not found")
    span = enclosing_table(x, pos)
    if not span:
        raise SystemExit("no enclosing table")
    tbl = x[span[0]:span[1]]
    sect = re.search(r'<w:sectPr>.*?</w:sectPr>', x, re.S)
    sectxml = sect.group(0) if sect else (
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>'
        '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
    doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document {NS}><w:body>{tbl}<w:p/>{sectxml}</w:body></w:document>')
    suffix = f"_probe{probe}" if probe else "_faithful"
    out = os.path.join(OUT, f"{doc_id}_J{suffix}.docx")
    os.makedirs(OUT, exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        zo.writestr('[Content_Types].xml', CT)
        zo.writestr('_rels/.rels', RELS)
        zo.writestr('word/_rels/document.xml.rels', DOCRELS)
        zo.writestr('word/styles.xml', styles)
        zo.writestr('word/document.xml', doc)
    # validate well-formed
    import xml.etree.ElementTree as ET
    try:
        ET.fromstring(doc)
        wf = "well-formed"
    except Exception as e:
        wf = f"XML ERROR: {e}"
    print(f"wrote {out}\n  table span {span} ({span[1]-span[0]} chars), tr={tbl.count('<w:tr')}, {wf}")
    return out


if __name__ == "__main__":
    if len(sys.argv) >= 4 and sys.argv[1] == "extract":
        extract(sys.argv[2], sys.argv[3])
    else:
        print(__doc__)
