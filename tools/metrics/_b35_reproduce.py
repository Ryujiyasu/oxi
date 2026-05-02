"""Reproduce b35 para 1 conditions: linesAndChars grid lp=350, jc=center,
sz=22 (=11pt), centered title. Goal: identify which axis suppresses
the centering offset that synthetic 12pt grid showed.
"""
import os, sys, time, json, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import win32com.client

OUT = os.path.join(os.path.dirname(__file__), "output", "b35_reproduce.json")
os.makedirs(os.path.dirname(OUT), exist_ok=True)
FIX_DIR = os.path.join(os.path.dirname(__file__), "output", "b35_repro_fixtures")
os.makedirs(FIX_DIR, exist_ok=True)


def build(out_path, lp_tw, sz_hp, grid_type, jc):
    """Build minimal docx mimicking b35 properties."""
    sp_jc = f'<w:jc w:val="{jc}"/>' if jc else ''
    rfonts = '<w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/>'
    sz = f'<w:sz w:val="{sz_hp}"/>'

    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>{sp_jc}<w:rPr>{rfonts}{sz}</w:rPr></w:pPr>
      <w:r><w:rPr>{rfonts}{sz}</w:rPr><w:t>Test</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1418" w:right="1418" w:bottom="1134" w:left="1418" w:header="851" w:footer="992"/>
      <w:docGrid w:type="{grid_type}" w:linePitch="{lp_tw}"/>
    </w:sectPr>
  </w:body>
</w:document>'''

    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", doc_xml)


def measure(word, path):
    last_err = None
    for attempt in range(3):
        try:
            wdoc = word.Documents.Open(path); break
        except Exception as e:
            last_err = e
            time.sleep(2)
            try:
                while word.Documents.Count > 0: word.Documents(1).Close(False)
            except: pass
    else:
        raise last_err
    try:
        wdoc.Repaginate(); time.sleep(0.1)
        p = wdoc.Paragraphs(1).Range
        y_para = round(p.Information(6), 4)
        # top margin (page Y=0 + topMargin)
        sec = wdoc.Sections(1)
        top_margin = round(sec.PageSetup.TopMargin, 4)
        return {"y_para": y_para, "top_margin": top_margin, "offset": round(y_para - top_margin, 4)}
    finally:
        wdoc.Close(False)


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False; word.DisplayAlerts = False; time.sleep(2.0)

    # Variants isolating each axis from b35:
    # b35 = linesAndChars + lp=350 + jc=center + sz=22 (11pt MS Gothic)
    cases = [
        ("V0_baseline_synthetic_lp360_lines",  360, 24, "lines",         None),    # control
        ("V1_b35_lp350",                       350, 24, "lines",         None),    # change pitch
        ("V2_b35_linesAndChars_lp360",         360, 24, "linesAndChars", None),    # change grid type
        ("V3_b35_linesAndChars_lp350",         350, 24, "linesAndChars", None),    # both
        ("V4_b35_linesAndChars_lp350_jcCenter", 350, 24, "linesAndChars", "center"), # add jc
        ("V5_b35_linesAndChars_lp350_sz22",    350, 22, "linesAndChars", None),    # change sz to 11pt
        ("V6_b35_linesAndChars_lp350_sz22_center", 350, 22, "linesAndChars", "center"), # full b35
        ("V7_b35_lines_lp350_sz22",            350, 22, "lines",         None),    # vary grid type w/ sz22
    ]

    results = []
    try:
        for label, lp, sz, gt, jc in cases:
            path = os.path.join(FIX_DIR, f"{label}.docx")
            build(path, lp, sz, gt, jc)
            try:
                r = measure(word, path)
                r["label"] = label
                r["lp_tw"] = lp; r["sz_hp"] = sz; r["grid_type"] = gt; r["jc"] = jc
                results.append(r)
                print(f"  {label}: y={r['y_para']} offset={r['offset']:+.2f}pt")
            except Exception as e:
                print(f"  {label}: ERR {str(e)[:60]}")
    finally:
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        try: word.Quit()
        except: pass
        print(f"\nSaved {len(results)} records.")


if __name__ == "__main__":
    main()
