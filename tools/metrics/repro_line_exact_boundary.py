"""Minimal repro for 2ea81 line=exact boundary +2pt bug.

3-paragraph docx:
- A: 'AAA' MS Mincho 10.5pt, line=260tw exact (13pt)
- B: empty, MS Mincho 10.5pt, line=260tw exact (13pt)
- C: 'CCC' MS Mincho 10.5pt, line=300tw exact (15pt)

Word expected: A→B=13pt, B→C=13pt → A→C=26pt.
Oxi buggy (pre-fix): A→C=28pt (uses idx C's 300tw at idx B→C boundary).
"""
import io, json, os, subprocess, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP_DOCX = Path("pipeline_data") / "_repro_line_exact.docx"
TMP_JSON = Path("pipeline_data") / "_repro_line_exact_layout.json"
OXI_RENDERER = Path(r"C:/Users/ryuji/oxi-4/tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def build():
    rpr = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
    # Para A: line=260 exact, "AAA"
    pa = f'<w:p><w:pPr><w:spacing w:line="260" w:lineRule="exact"/>{rpr}</w:pPr><w:r>{rpr}<w:t>AAA</w:t></w:r></w:p>'
    # Para B: line=260 exact, empty
    pb = f'<w:p><w:pPr><w:spacing w:line="260" w:lineRule="exact"/>{rpr}</w:pPr></w:p>'
    # Para C: line=300 exact, "CCC"
    pc = f'<w:p><w:pPr><w:spacing w:line="300" w:lineRule="exact"/>{rpr}</w:pPr><w:r>{rpr}<w:t>CCC</w:t></w:r></w:p>'
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{pa}{pb}{pc}{sect}</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP_DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def measure_word():
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(str(TMP_DOCX.resolve()), ReadOnly=True)
        time.sleep(0.5)
        ya = doc.Paragraphs(1).Range.Information(6)
        yb = doc.Paragraphs(2).Range.Information(6)
        yc = doc.Paragraphs(3).Range.Information(6)
        doc.Close(False)
        return {"A": round(ya, 2), "B": round(yb, 2), "C": round(yc, 2)}
    finally:
        word.Quit()


def measure_oxi():
    try: os.remove(TMP_JSON)
    except FileNotFoundError: pass
    trash = str(Path("pipeline_data") / "_repro_trash").replace("\\", "/")
    result = subprocess.run([str(OXI_RENDERER), str(TMP_DOCX.resolve()), trash, "150", f"--dump-layout={TMP_JSON.resolve()}"],
                            capture_output=True, timeout=30)
    if result.returncode != 0:
        print(f"renderer err: {result.stderr.decode('utf-8', errors='replace')[:300]}")
        return None
    d = json.load(open(TMP_JSON, encoding='utf-8'))
    p1 = d['pages'][0]['elements']
    # Find A, B, C text positions
    a_y = b_y = c_y = None
    for e in p1:
        if e.get('type') != 'text': continue
        t = e.get('text', '') or ''
        if t.startswith('A') and a_y is None: a_y = e.get('y')
        elif t.startswith('C') and c_y is None: c_y = e.get('y')
    # B is empty — infer from A→C distance vs pitch
    return {"A": a_y, "C": c_y}


build()
print(f"Built {TMP_DOCX}")
w = measure_word()
print(f"Word: A={w['A']} B={w['B']} C={w['C']}  A→B={w['B']-w['A']:.2f}  B→C={w['C']-w['B']:.2f}  A→C={w['C']-w['A']:.2f}")
o = measure_oxi()
if o:
    print(f"Oxi:  A={o['A']} C={o['C']}  A→C={o['C']-o['A']:.2f}" if o['A'] and o['C'] else f"Oxi: {o}")
    delta = (o['C'] - o['A']) - (w['C'] - w['A']) if o['A'] and o['C'] else None
    print(f"  → Oxi A→C - Word A→C = {delta:+.2f}pt")
