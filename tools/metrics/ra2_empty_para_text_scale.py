"""Test if `<w:w val="N"/>` affects empty-paragraph line height in Word.

3a4f9f has many empty paragraphs with pPr.rPr structure:
  <w:rPr><w:w w:val="200"/><w:sz w:val="48"/></w:rPr>

If Word treats text_scale as affecting empty-para LH, Oxi's current
implementation (which only applies text_scale to char widths) would
under/over-estimate empty para LH.

Test: 5-paragraph fixture with empty paragraphs at sz=24pt, varying text_scale.
"""
import os
import sys
import time
import json

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "empty_para_text_scale_fixtures")
os.makedirs(FIX_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_empty_para_text_scale.json")

WD_LAYOUT_LINEGRID = 2


def build_via_ooxml(out_path, scale_val):
    """Build minimal docx with 5 paragraphs:
    P1: text "Top"
    P2: empty pPr.rPr.w=scale_val sz=48 (24pt)
    P3: empty same
    P4: empty same
    P5: text "Bot"
    """
    import zipfile
    from io import BytesIO

    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>Top</w:t></w:r></w:p>
    <w:p><w:pPr><w:rPr><w:w w:val="{scale_val}"/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr></w:pPr></w:p>
    <w:p><w:pPr><w:rPr><w:w w:val="{scale_val}"/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr></w:pPr></w:p>
    <w:p><w:pPr><w:rPr><w:w w:val="{scale_val}"/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr></w:pPr></w:p>
    <w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>Bot</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:docGrid w:type="lines" w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>'''

    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", document_xml)


def measure(word, path):
    last_err = None
    for attempt in range(3):
        try:
            wdoc = word.Documents.Open(path)
            break
        except Exception as e:
            last_err = e
            time.sleep(2)
            try:
                while word.Documents.Count > 0:
                    word.Documents(1).Close(False)
            except Exception:
                pass
    else:
        raise last_err
    try:
        wdoc.Repaginate()
        time.sleep(0.1)
        ys = []
        for i in range(1, wdoc.Paragraphs.Count + 1):
            p = wdoc.Paragraphs(i)
            ys.append(round(p.Range.Information(6), 4))
        return ys
    finally:
        wdoc.Close(False)


def main():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)

    results = {}
    try:
        for scale in [100, 150, 200, 50]:
            path = os.path.join(FIX_DIR, f"empty_w{scale}.docx")
            build_via_ooxml(path, scale)
            try:
                ys = measure(word, path)
                gaps = [round(ys[i+1] - ys[i], 4) for i in range(len(ys) - 1)]
                results[scale] = {"ys": ys, "gaps": gaps}
                print(f"  w_val={scale}: ys={ys}")
                print(f"    gaps (P2-P1, P3-P2, P4-P3, P5-P4): {gaps}")
            except Exception as e:
                print(f"  w_val={scale}: ERR {e}")
                results[scale] = {"error": str(e)}
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved to {OUT_JSON}")
        try: word.Quit()
        except: pass


if __name__ == "__main__":
    main()
