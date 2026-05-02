"""§5.DD round 81 — evenPage/oddPage section break blank page insertion.

ECMA-376 §17.6.21 — section break types evenPage/oddPage force
the next section to start on an even/odd page, inserting a blank
page if necessary.

Probes (P1 on page 1 = odd):
  V1 type=nextPage: P2 on page 2 (next page, no blank)
  V2 type=evenPage: P2 on page 2 (already even after page 1, no blank needed)
  V3 type=oddPage: P2 on page 3 (blank page 2 inserted)
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\section_break_evenodd_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\section_break_evenodd.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(section_break_type, margin_tw, page_w_tw):
    """Two paragraphs in 2 sections separated by break of given type."""
    p1 = ('<w:p><w:pPr>'
          '<w:jc w:val="left"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
          # Inline sectPr inside the paragraph
          '<w:sectPr>'
          f'<w:type w:val="{section_break_type}"/>'
          f'<w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
          f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
          ' w:header="720" w:footer="720" w:gutter="0"/>'
          '<w:cols w:space="720"/>'
          '<w:docGrid w:linePitch="360"/>'
          '</w:sectPr>'
          '</w:pPr>'
          '<w:r><w:rPr>'
          f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
          '<w:sz w:val="24"/></w:rPr>'
          '<w:t>S1Para</w:t></w:r></w:p>')
    p2 = ('<w:p><w:pPr><w:jc w:val="left"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
          '</w:pPr>'
          '<w:r><w:rPr>'
          f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
          '<w:sz w:val="24"/></w:rPr>'
          '<w:t>S2Para</w:t></w:r></w:p>')
    body = (f'{p1}{p2}'
            '<w:sectPr>'
            f'<w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
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
                    page = int(rng.Information(3))
                    text = rng.Text[:30] if rng.Text else ""
                    para_info.append({"i": pi, "page": page, "text_preview": text})
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

    print(f"evenPage/oddPage section break test\n")

    variants = [
        ("V1_nextPage", "nextPage", "P2 on next page (=2, even)"),
        ("V2_evenPage", "evenPage", "P2 on next even page (=2)"),
        ("V3_oddPage",  "oddPage",  "P2 on next odd page (=3, blank page 2)"),
    ]

    for label, brk_type, desc in variants:
        doc_xml = make_doc_xml(brk_type, margin_tw, page_w_tw)
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
            print(f"      para[{p['i']}]: page={p['page']} text={p['text_preview']!r}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
