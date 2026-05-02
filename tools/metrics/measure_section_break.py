"""§5.X round 69 — section break types.

ECMA-376 §17.6.21 — `<w:sectPr><w:type w:val="..."/></w:sectPr>`:
  - continuous: same page, new section
  - nextPage (default): new page
  - nextColumn: new column (multi-col)
  - evenPage: next even-numbered page
  - oddPage: next odd-numbered page

Probes: 2 paragraphs in 2 sections separated by various break types.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\section_break_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\section_break.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(section_break_type, page_w_tw, margin_tw):
    """Two sections separated by break of given type."""
    type_attr = f'<w:type w:val="{section_break_type}"/>' if section_break_type else ''
    p1 = ('<w:p><w:pPr>'
          '<w:jc w:val="left"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
          # Inline sectPr inside the paragraph for first section
          '<w:sectPr>'
          f'{type_attr}'
          f'<w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
          f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}" w:header="720" w:footer="720" w:gutter="0"/>'
          '<w:cols w:space="720"/>'
          '<w:docGrid w:linePitch="360"/>'
          '</w:sectPr>'
          '</w:pPr>'
          '<w:r><w:rPr>'
          f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
          '<w:sz w:val="24"/></w:rPr>'
          '<w:t>S1Para1</w:t></w:r></w:p>')
    p2 = ('<w:p><w:pPr><w:jc w:val="left"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
          '</w:pPr>'
          '<w:r><w:rPr>'
          f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
          '<w:sz w:val="24"/></w:rPr>'
          '<w:t>S2Para1</w:t></w:r></w:p>')
    # End-of-doc sectPr (for section 2)
    body = (f'{p1}{p2}'
            f'<w:sectPr>'
            f'<w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}" w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}</w:body></w:document>')


def make_styles_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            '<w:sz w:val="24"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>')


def make_docx(label, doc_xml):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    styles_xml = make_styles_xml()
    settings_xml = make_settings_xml()
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '</Types>'
    )
    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>'
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
        ' Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
        ' Target="settings.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", styles_xml)
    return out_path


def kill_word():
    subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
    time.sleep(2)


def measure_one(path):
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        d = word.Documents.Open(str(path), ReadOnly=True)
        time.sleep(0.3)
        try:
            n_paras = d.Paragraphs.Count
            para_info = []
            for pi in range(1, n_paras + 1):
                p = d.Paragraphs(pi)
                rng = p.Range
                try:
                    y = float(rng.Information(6))
                    page = int(rng.Information(3))
                    text = rng.Text[:30] if rng.Text else ""
                    para_info.append({"i": pi, "y": y, "page": page, "text_preview": text})
                except: continue
            n_pages = int(d.Range().Information(4))  # wdNumberOfPagesInDocument
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        return {"n_paras": len(para_info), "n_pages": n_pages, "paragraphs": para_info}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    margin_tw = 170 * 10
    page_w_tw = int((400 + 170) * 20)

    print(f"Section break test, fs=12 MS Mincho\n")

    variants = [
        ("V1_continuous", "continuous", "section break continuous (same page)"),
        ("V2_nextPage",   "nextPage",   "section break nextPage (new page)"),
        ("V3_evenPage",   "evenPage",   "section break evenPage"),
        ("V4_oddPage",    "oddPage",    "section break oddPage"),
        ("V5_default",    None,         "default (no type, equiv to nextPage)"),
    ]

    for label, brk_type, desc in variants:
        doc_xml = make_doc_xml(brk_type, page_w_tw, margin_tw)
        try:
            p = make_docx(label, doc_xml)
        except Exception as e:
            out[label] = {"build_error": str(e)}
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        entry = {"label": label, "desc": desc, **r}
        out[label] = entry
        print(f"  {label}: {desc}")
        print(f"    n_pages={entry.get('n_pages')} n_paras={entry.get('n_paras')}")
        for p in entry.get("paragraphs", []):
            print(f"      para[{p['i']}]: page={p['page']} y={p['y']} text={p['text_preview']!r}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
