"""§5.CC round 66 — w:pageBreakBefore paragraph property.

ECMA-376 §17.3.1.27 — `<w:pageBreakBefore/>` in pPr causes the
paragraph to start on a new page.

Different from `<w:br w:type="page"/>` (Round 65): pageBreakBefore
is paragraph-level (cleaner) vs br-type-page is run-level (uses
form feed char).

Probes:
  V1 plain (no break)
  V2 P2 with pageBreakBefore
  V3 first paragraph with pageBreakBefore (should it still break?)
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\page_break_before_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\page_break_before.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_para(text, sz_val, page_break_before=False):
    pbb = '<w:pageBreakBefore/>' if page_break_before else ''
    return ('<w:p><w:pPr>'
            f'{pbb}'
            '<w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>')


def make_doc_xml(body_inner_xml, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body_inner_xml}'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml(sz_val):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>')


def make_docx(label, doc_xml, sz_val):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    styles_xml = make_styles_xml(sz_val)
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
                    page_num = int(rng.Information(3))
                    text = rng.Text[:30] if rng.Text else ""
                    para_info.append({"i": pi, "y": y, "page": page_num, "text_preview": text})
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        return {"n_paras": len(para_info), "paragraphs": para_info}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    margin_tw = 170 * 10
    page_w_tw = int((400 + 170) * 20)

    # V1: 2 plain paragraphs
    body_v1 = make_para("Page1A", sz_val) + make_para("Page2A", sz_val)

    # V2: P2 has pageBreakBefore
    body_v2 = make_para("Page1A", sz_val) + make_para("Page2A", sz_val, page_break_before=True)

    # V3: P1 has pageBreakBefore
    body_v3 = make_para("Page1A", sz_val, page_break_before=True) + make_para("Page2A", sz_val)

    # V4: 3 paras, P2 has pageBreakBefore
    body_v4 = (make_para("Para1", sz_val)
               + make_para("Para2", sz_val, page_break_before=True)
               + make_para("Para3", sz_val))

    print(f"pageBreakBefore test, fs=12 MS Mincho\n")

    variants = [
        ("V1_plain", body_v1, "2 plain paragraphs (control)"),
        ("V2_p2_break", body_v2, "P2 has pageBreakBefore"),
        ("V3_p1_break", body_v3, "P1 (first) has pageBreakBefore"),
        ("V4_three_p2_break", body_v4, "3 paras, P2 has pageBreakBefore"),
    ]

    for label, body_xml, desc in variants:
        doc_xml = make_doc_xml(body_xml, page_w_tw, margin_tw)
        try:
            p = make_docx(label, doc_xml, sz_val)
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
        for p in entry.get("paragraphs", []):
            print(f"    para[{p['i']}]: page={p['page']} y={p['y']} text={p['text_preview']!r}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
