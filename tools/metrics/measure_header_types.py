"""§5.DD round 83 — first/odd/even header types.

ECMA-376:
  §17.6.7 titlePg — section property to enable first-page header
  §17.10.6 evenAndOddHeaders — settings.xml for separate odd/even
  §17.10.x headerReference w:type — default | first | even

Probes:
  V1 default header only (all pages same)
  V2 default + first (titlePg set, first page uses first header)
  V3 default + even (evenAndOddHeaders set, page 2 uses even header)
  V4 default + first + even (all 3 types in 2-page doc)
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\header_types_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\header_types.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_header_xml(text):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:p><w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            '<w:sz w:val="24"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p></w:hdr>')


def make_doc_xml(margin_tw, page_w_tw, headers, title_pg=False):
    """headers: dict of type → r:id (e.g., {"default": "rId3", "first": "rId4"})."""
    body_inner = ('<w:p><w:pPr><w:jc w:val="left"/>'
                  '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
                  '</w:pPr>'
                  '<w:r><w:rPr>'
                  f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
                  '<w:sz w:val="24"/></w:rPr>'
                  '<w:t>P1</w:t>'
                  '<w:br w:type="page"/>'
                  '<w:t>P2</w:t></w:r></w:p>')
    header_refs = "".join(f'<w:headerReference w:type="{t}" r:id="{rid}"/>'
                          for t, rid in headers.items())
    title_pg_xml = '<w:titlePg/>' if title_pg else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<w:body>{body_inner}'
            '<w:sectPr>'
            f'{header_refs}'
            f'{title_pg_xml}'
            f'<w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


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


def make_settings_xml(even_and_odd=False):
    inner = '<w:evenAndOddHeaders/>' if even_and_odd else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'{inner}'
            '</w:settings>')


def make_docx(label, doc_xml, settings_xml, header_files):
    """header_files: dict of relationship_id → (filename, xml_content)."""
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    styles_xml = make_styles_xml()
    # Build content_types
    overrides = (
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>')
    for rid, (fname, _) in header_files.items():
        overrides += (f'<Override PartName="/word/{fname}"'
                      ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>')
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'{overrides}'
        '</Types>')
    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>')
    # Build doc.xml.rels
    rels_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
                ' Target="styles.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
                ' Target="settings.xml"/>')
    for rid, (fname, _) in header_files.items():
        rels_xml += (f'<Relationship Id="{rid}"'
                     ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"'
                     f' Target="{fname}"/>')
    rels_xml += '</Relationships>'

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", rels_xml)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", styles_xml)
        z.writestr("word/settings.xml", settings_xml)
        for rid, (fname, content) in header_files.items():
            z.writestr(f"word/{fname}", content)
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
            # Iterate through Headers
            sections_data = []
            for sec in d.Sections:
                sec_data = {}
                # wdHeaderFooterPrimary=1, wdHeaderFooterFirstPage=2, wdHeaderFooterEvenPages=3
                names = ["primary", "first", "even"]
                for idx, hf in enumerate(sec.Headers):
                    if idx >= 3: break
                    try:
                        sec_data[names[idx]] = hf.Range.Text[:30]
                    except: sec_data[names[idx]] = "<error>"
                sections_data.append(sec_data)
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        return {"sections": sections_data}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    margin_tw = 1700
    page_w_tw = int((400 + 170) * 20)

    print(f"Header types test\n")

    # V1 default only
    body_v1 = make_doc_xml(margin_tw, page_w_tw, {"default": "rId3"})
    files_v1 = {"rId3": ("header_default.xml", make_header_xml("DefaultHdr"))}

    # V2 default + first + titlePg
    body_v2 = make_doc_xml(margin_tw, page_w_tw,
                           {"default": "rId3", "first": "rId4"},
                           title_pg=True)
    files_v2 = {"rId3": ("header_default.xml", make_header_xml("DefaultHdr")),
                "rId4": ("header_first.xml", make_header_xml("FirstHdr"))}

    # V3 default + even + evenAndOddHeaders setting
    body_v3 = make_doc_xml(margin_tw, page_w_tw,
                           {"default": "rId3", "even": "rId4"})
    files_v3 = {"rId3": ("header_default.xml", make_header_xml("OddHdr")),
                "rId4": ("header_even.xml", make_header_xml("EvenHdr"))}

    variants = [
        ("V1_default_only", body_v1, make_settings_xml(False), files_v1, "default header only"),
        ("V2_default_first", body_v2, make_settings_xml(False), files_v2, "default + first (titlePg)"),
        ("V3_default_even", body_v3, make_settings_xml(True), files_v3, "default + even (evenAndOddHeaders)"),
    ]

    for label, doc_xml, settings_xml, files, desc in variants:
        try:
            p = make_docx(label, doc_xml, settings_xml, files)
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
        for sec in entry.get("sections", []):
            print(f"    section headers: {sec}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
