"""§5.X round 76 — per-section page margins.

ECMA-376 §17.6.11 — `<w:pgMar>` per section. Test 2 sections
with different margins.

Probes:
  V1 single section margin=85pt: x=85
  V2 2 sections, S1=85pt + S2=120pt
  V3 2 sections, S1=120pt + S2=85pt
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\section_margins_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\section_margins.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_section_props(margin_l_tw, margin_r_tw, page_w_tw):
    return ('<w:sectPr>'
            f'<w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_r_tw}" w:bottom="1134" w:left="{margin_l_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr>')


def make_para_with_inline_sectpr(text, sect_xml, sz_val):
    return ('<w:p><w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'{sect_xml}'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>')


def make_para(text, sz_val):
    return ('<w:p><w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>')


def make_doc_xml(body_inner_xml):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body_inner_xml}</w:body></w:document>')


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
            chars = d.Range().Characters
            xs = []
            for ci in range(1, chars.Count + 1):
                try:
                    c = chars(ci)
                    t = c.Text
                    if not t or any(ord(ch) < 32 for ch in t): continue
                    xs.append((t, float(c.Information(5)), float(c.Information(6)), int(c.Information(3))))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        return {"n_chars": len(xs),
                "first_5": [{"ch": xs[i][0], "x": xs[i][1], "y": xs[i][2], "page": xs[i][3]}
                            for i in range(min(5, len(xs)))]}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    page_w_tw = 8500  # ~425pt page wide

    # Margin variants in twips
    m85 = 1700  # 85pt
    m120 = 2400  # 120pt

    print(f"Per-section page margin test, fs=12 MS Mincho\n")

    # V1 single section margin=85pt
    body_v1 = make_para("S1Para", sz_val) + make_section_props(m85, m85, page_w_tw)

    # V2 S1=85, S2=120 (continuous to keep on same page... but Round 69 showed all break)
    body_v2 = (make_para_with_inline_sectpr("S1Para",
               '<w:sectPr>'
               f'<w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
               f'<w:pgMar w:top="1134" w:right="{m85}" w:bottom="1134" w:left="{m85}"'
               ' w:header="720" w:footer="720" w:gutter="0"/>'
               '<w:cols w:space="720"/>'
               '<w:docGrid w:linePitch="360"/>'
               '</w:sectPr>', sz_val)
               + make_para("S2Para", sz_val)
               + make_section_props(m120, m120, page_w_tw))

    # V3 reverse: S1=120, S2=85
    body_v3 = (make_para_with_inline_sectpr("S1Para",
               '<w:sectPr>'
               f'<w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
               f'<w:pgMar w:top="1134" w:right="{m120}" w:bottom="1134" w:left="{m120}"'
               ' w:header="720" w:footer="720" w:gutter="0"/>'
               '<w:cols w:space="720"/>'
               '<w:docGrid w:linePitch="360"/>'
               '</w:sectPr>', sz_val)
               + make_para("S2Para", sz_val)
               + make_section_props(m85, m85, page_w_tw))

    variants = [
        ("V1_single_85", body_v1, "single section margin=85pt"),
        ("V2_S1_85_S2_120", body_v2, "S1 margin=85, S2 margin=120"),
        ("V3_S1_120_S2_85", body_v3, "S1 margin=120, S2 margin=85"),
    ]

    for label, body_xml, desc in variants:
        doc_xml = make_doc_xml(body_xml)
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
        chars = entry.get("first_5", [])
        print(f"  {label}: {desc}")
        for c in chars:
            print(f"    {c['ch']!r}: page={c['page']} x={c['x']} y={c['y']}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
