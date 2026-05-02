# -*- coding: utf-8 -*-
"""V_II: Minimal PAGE field in footer test. Measure Word's rendering.

Variants test:
- V_II0: simple PAGE in centered footer
- V_II1: PAGE/NUMPAGES "1/N" pattern
- V_II2: 2-page doc — page number changes per page
- V_II3: Field with formatting (\* MERGEFORMAT)
- V_II4: PAGE in multi-run with leading text "Page "
"""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client as wc
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES as CTYPES_BASE, RELS_ROOT

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/v_ii_page_field")
os.makedirs(OUT_DIR, exist_ok=True)

# Need Content_Types with footer override
CTYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>'''

WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:kern w:val="2"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:defaultTabStop w:val="840"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>
</w:settings>'''


# Footer XML variants
def footer_xml_v0():
    """Simple: <PAGE> in centered paragraph"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p>
<w:pPr><w:jc w:val="center"/></w:pPr>
<w:fldSimple w:instr="PAGE">
<w:r><w:t>1</w:t></w:r>
</w:fldSimple>
</w:p>
</w:ftr>'''


def footer_xml_v1():
    """Page/NumPages: 1 / N"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p>
<w:pPr><w:jc w:val="center"/></w:pPr>
<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>
<w:r><w:t xml:space="preserve"> / </w:t></w:r>
<w:fldSimple w:instr="NUMPAGES"><w:r><w:t>1</w:t></w:r></w:fldSimple>
</w:p>
</w:ftr>'''


def footer_xml_v3():
    """fldChar style with MERGEFORMAT"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p>
<w:pPr><w:jc w:val="center"/></w:pPr>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
</w:ftr>'''


def footer_xml_v4():
    """Leading text + PAGE: "Page 1" """
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p>
<w:pPr><w:jc w:val="center"/></w:pPr>
<w:r><w:t xml:space="preserve">Page </w:t></w:r>
<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>
</w:p>
</w:ftr>'''


def doc_xml(*, body_pages):
    """body_pages: number of pages to force (via page breaks)"""
    paras = ['<w:p><w:r><w:t>Page 1 body</w:t></w:r></w:p>']
    for i in range(2, body_pages + 1):
        paras.append(f'<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>Page {i} body</w:t></w:r></w:p>')
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
{"".join(paras)}
<w:sectPr>
<w:footerReference w:type="default" r:id="rId3"/>
<w:pgSz w:w="11906" w:h="16838" w:code="9"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>'''


def write_docx(path, *, footer_content, body_pages=1):
    tmp = tempfile.mkdtemp(prefix='ii_')
    try:
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
            ("word/footer1.xml", footer_content),
            ("word/document.xml", doc_xml(body_pages=body_pages)),
        ]
        for relpath, content in files:
            full = os.path.join(tmp, relpath.replace("/", os.sep))
            os.makedirs(os.path.dirname(full), exist_ok=True)
            with open(full, "w", encoding="utf-8") as f:
                f.write(content)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace(os.sep, '/')
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def render_pdf(word, docx, pdf):
    last = None
    for attempt in range(5):
        try:
            d = word.Documents.Open(docx, ReadOnly=True)
            time.sleep(0.4)
            d.SaveAs2(pdf, FileFormat=17)
            d.Close(SaveChanges=False)
            return True
        except Exception as e:
            last = e
            time.sleep(1.0 + attempt * 0.5)
    print(f"  ERR: {last}")
    return False


def measure_footer_text(pdf):
    """Find numbers and expected page indicators in PDF (likely in footer area)."""
    d = fitz.open(pdf)
    res = []
    for pi in range(d.page_count):
        page = d[pi]
        # Find all text in lower 100pt of page (footer area)
        page_h = page.rect.height
        footer_area = fitz.Rect(0, page_h - 80, page.rect.width, page_h)
        txt = page.get_text("text", clip=footer_area)
        # Also get individual instances of "1", "2", "/", "Page" etc.
        items = []
        for ch in ['1', '2', '3', '/', 'Page']:
            for inst in page.search_for(ch, clip=footer_area):
                items.append({'char': ch, 'x': inst.x0, 'y': inst.y0})
        res.append({'page': pi+1, 'footer_text': txt.strip(), 'positions': items})
    d.close()
    return res


VARIANTS = [
    ("V_II0_simple_PAGE", {"footer_content": footer_xml_v0(), "body_pages": 1}),
    ("V_II1_page_numpages_2pgs", {"footer_content": footer_xml_v1(), "body_pages": 2}),
    ("V_II2_2pages_PAGE_only", {"footer_content": footer_xml_v0(), "body_pages": 2}),
    ("V_II3_fldChar_PAGE_2pgs", {"footer_content": footer_xml_v3(), "body_pages": 2}),
    ("V_II4_text_plus_PAGE", {"footer_content": footer_xml_v4(), "body_pages": 2}),
]


def main():
    pythoncom.CoInitialize()
    word = None
    for attempt in range(5):
        try:
            word = wc.Dispatch("Word.Application")
            time.sleep(2.0)
            word.Visible = False
            word.DisplayAlerts = False
            break
        except Exception as e:
            print(f"Word startup {attempt+1}: {e}")
            time.sleep(8.0)
    if word is None:
        print("Failed Word"); return
    print("V_II: PAGE field in footer test\n")
    results = []
    try:
        for vid, kwargs in VARIANTS:
            print(f"=== {vid} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            write_docx(docx, **kwargs)
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render"})
                continue
            footer_data = measure_footer_text(pdf)
            for fd in footer_data:
                print(f"  Page {fd['page']}: footer text = {fd['footer_text']!r}")
                for item in fd['positions']:
                    print(f"    '{item['char']}' at x={item['x']:.2f} y={item['y']:.2f}")
            results.append({"id": vid, "footer_data": footer_data})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
