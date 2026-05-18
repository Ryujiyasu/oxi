"""Focused probe: isolate cell.mar_top behavior.

Build a series of minimal docs varying ONLY cell.mar_top while keeping
everything else fixed. Measure Word vs Oxi and look for tcMar
interpretation divergence.

This addresses the alpha01 fuzz signal where cell.mar_top dominated
top-15 attribute pairs.
"""
import json, os, subprocess, sys, tempfile, zipfile
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
OUT_DIR = ROOT / "tools/metrics/fuzz_runs/tcmar_probe"
OUT_DIR.mkdir(parents=True, exist_ok=True)
RENDERER = ROOT / "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe"


def make_docx(name: str, mar_top: int | None, mar_bottom: int | None = None):
    """Single-row table with explicit tcMar.top variation."""
    mar_parts = []
    if mar_top is not None:
        mar_parts.append(f'<w:top w:w="{mar_top}" w:type="dxa"/>')
    if mar_bottom is not None:
        mar_parts.append(f'<w:bottom w:w="{mar_bottom}" w:type="dxa"/>')
    mar_xml = f'<w:tcMar>{"".join(mar_parts)}</w:tcMar>' if mar_parts else ""

    body = f'''<w:p><w:r><w:t>HEAD</w:t></w:r></w:p>
<w:tbl>
<w:tblPr>
<w:tblW w:w="6000" w:type="dxa"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>
<w:tr><w:tc>
<w:tcPr><w:tcW w:w="6000" w:type="dxa"/>{mar_xml}</w:tcPr>
<w:p><w:r><w:t>サンプル</w:t></w:r></w:p>
</w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>ANCHOR</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701"/>
<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/></w:sectPr>'''

    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>{body}</w:body></w:document>'''

    CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>'''
    out = OUT_DIR / f"{name}.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc_xml)
    return out


def collapse_y(rng):
    doc = rng.Document
    return doc.Range(rng.Start, rng.Start).Information(6)


def measure(name: str, mar_top: int | None):
    docx = make_docx(name, mar_top)
    # Word
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx.absolute()), ReadOnly=True)
        try:
            head_y = None; cell_y = None; anchor_y = None
            for p in doc.Paragraphs:
                t = p.Range.Text.strip()
                if t == "HEAD":
                    head_y = collapse_y(p.Range)
                elif t == "ANCHOR":
                    anchor_y = collapse_y(p.Range)
            cell = doc.Tables(1).Cell(Row=1, Column=1)
            cell_y = collapse_y(cell.Range)
            w_data = {"head": head_y, "cell": cell_y, "anchor": anchor_y}
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    # Oxi
    with tempfile.TemporaryDirectory() as tmp:
        prefix = os.path.join(tmp, "p_")
        dump = os.path.join(tmp, "layout.json")
        subprocess.run([str(RENDERER), str(docx), prefix, "--dump-layout="+dump],
                      capture_output=True, text=True, timeout=60)
        with open(dump, encoding="utf-8") as f:
            d = json.load(f)
    p0 = d["pages"][0]
    o_head = o_cell = o_anchor = None
    for el in p0["elements"]:
        if el.get("type") != "text": continue
        text = el.get("text", "")
        if el.get("cell_row_idx") == 0 and el.get("cell_col_idx") == 0:
            o_cell = el["y"]
        elif "HEAD" in text:
            o_head = el["y"]
        elif "ANCHOR" in text:
            o_anchor = el["y"]
    o_data = {"head": o_head, "cell": o_cell, "anchor": o_anchor}
    return w_data, o_data


def main():
    variants = [None, 0, 12, 50, 100, 200, 500]
    print(f"{'mar_top':<10} {'W_head':<8} {'W_cell':<8} {'W_anchor':<10} {'O_head':<8} {'O_cell':<8} {'O_anchor':<10} {'cell_diff':<10}")
    results = []
    for mt in variants:
        w, o = measure(f"tcmar_{str(mt)}", mt)
        cell_diff = (w["cell"] - w["head"]) - (o["cell"] - o["head"])
        anchor_diff = (w["anchor"] - w["head"]) - (o["anchor"] - o["head"])
        results.append({"mar_top": mt, "word": w, "oxi": o, "cell_dy_diff": cell_diff, "anchor_dy_diff": anchor_diff})
        print(f"{str(mt):<10} {w['head']:<8.2f} {w['cell']:<8.2f} {w['anchor']:<10.2f} {o['head']:<8.2f} {o['cell']:<8.2f} {o['anchor']:<10.2f} {cell_diff:+.2f}")
    (OUT_DIR / "results.json").write_text(json.dumps(results, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
