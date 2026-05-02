"""§5.DD round 77 — per-section page size.

ECMA-376 §17.6.7 — `<w:pgSz>` per section. Test 2 sections with
different page widths and orientations.

Probes:
  V1 single section: pgSz w=8500 (425pt)
  V2 S1 portrait + S2 landscape (w/h swapped)
  V3 S1 small + S2 large
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\section_page_size_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\section_page_size.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_section_props(margin_l_tw, page_w_tw, page_h_tw, orient=None):
    orient_attr = f' w:orient="{orient}"' if orient else ''
    return ('<w:sectPr>'
            f'<w:pgSz w:w="{page_w_tw}" w:h="{page_h_tw}"{orient_attr}/>'
            f'<w:pgMar w:top="1134" w:right="{margin_l_tw}" w:bottom="1134" w:left="{margin_l_tw}"'
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
            # Page width info
            page_setup = d.PageSetup
            ps_pages = []
            try:
                # Word's PageSetup is per-section, so iterate sections
                for sec in d.Sections:
                    ps = sec.PageSetup
                    ps_pages.append({"page_w_pt": float(ps.PageWidth),
                                     "page_h_pt": float(ps.PageHeight),
                                     "orient": int(ps.Orientation)})
            except: pass
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars", "sections": ps_pages}
        return {"n_chars": len(xs),
                "first_5": [{"ch": xs[i][0], "x": xs[i][1], "y": xs[i][2], "page": xs[i][3]}
                            for i in range(min(5, len(xs)))],
                "last_3": [{"ch": xs[i][0], "x": xs[i][1], "y": xs[i][2], "page": xs[i][3]}
                            for i in range(max(0, len(xs)-3), len(xs))],
                "sections": ps_pages}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    margin_tw = 1700  # 85pt margins for both sections

    # V1 single section, w=8500 (425pt), portrait
    body_v1 = make_para("S1A", sz_val) + make_section_props(margin_tw, 8500, 16838)

    # V2 S1 portrait (w=8500 h=16838) + S2 landscape (w=16838 h=8500)
    body_v2 = (make_para_with_inline_sectpr("S1A",
               make_section_props(margin_tw, 8500, 16838).replace('<w:sectPr>', '<w:sectPr>').replace('</w:sectPr>', '</w:sectPr>'),
               sz_val)
               + make_para("S2A", sz_val)
               + make_section_props(margin_tw, 16838, 8500, orient="landscape"))

    # V3 S1 small (w=6000) + S2 large (w=12000)
    body_v3 = (make_para_with_inline_sectpr("S1A",
               make_section_props(margin_tw, 6000, 16838),
               sz_val)
               + make_para("S2A", sz_val)
               + make_section_props(margin_tw, 12000, 16838))

    print(f"Per-section page size test, fs=12 MS Mincho\n")

    variants = [
        ("V1_single", body_v1, "single section pgSz=8500x16838 (portrait)"),
        ("V2_portrait_landscape", body_v2, "S1 portrait + S2 landscape"),
        ("V3_small_large", body_v3, "S1 w=6000 + S2 w=12000"),
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
        print(f"  {label}: {desc}")
        print(f"    sections: {entry.get('sections')}")
        for c in entry.get("first_5", [])[:2]:
            print(f"      first: {c['ch']!r} page={c['page']} x={c['x']} y={c['y']}")
        for c in entry.get("last_3", []):
            print(f"      last: {c['ch']!r} page={c['page']} x={c['x']} y={c['y']}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
