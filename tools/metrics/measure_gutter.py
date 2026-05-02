"""§5.DD round 79 — w:gutter binding margin.

ECMA-376 §17.6.11 — pgMar/@gutter adds extra margin for binding.

Probes:
  V1 no gutter (gutter=0): reference
  V2 gutter=720tw (36pt): text shifts right by gutter amount
  V3 gutter=720 + mirror: alternate sides per page
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\gutter_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\gutter.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(margin_l_tw, margin_r_tw, gutter_tw):
    p1 = ('<w:p><w:pPr><w:jc w:val="left"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
          '</w:pPr>'
          '<w:r><w:rPr>'
          f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
          '<w:sz w:val="24"/></w:rPr>'
          '<w:t>P1A</w:t>'
          '<w:br w:type="page"/>'
          '<w:t>P2A</w:t></w:r></w:p>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{p1}'
            '<w:sectPr><w:pgSz w:w="12000" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_r_tw}" w:bottom="1134" w:left="{margin_l_tw}"'
            f' w:header="720" w:footer="720" w:gutter="{gutter_tw}"/>'
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


def make_settings_xml(mirror=False):
    inner = '<w:mirrorMargins/>' if mirror else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'{inner}'
            '</w:settings>')


def make_docx(label, doc_xml, settings_xml):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    styles_xml = make_styles_xml()
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
        z.writestr("word/settings.xml", settings_xml)
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
                    xs.append((t, float(c.Information(5)), int(c.Information(3))))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        page1 = [x for x in xs if x[2] == 1]
        page2 = [x for x in xs if x[2] == 2]
        return {"n_chars": len(xs),
                "page1_first_x": page1[0][1] if page1 else None,
                "page2_first_x": page2[0][1] if page2 else None}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    margin_l_tw = 1700  # 85pt
    margin_r_tw = 1700  # 85pt (symmetric)

    print(f"Gutter test: L=85pt, R=85pt\n")

    variants = [
        ("V1_no_gutter", 0, False, "no gutter"),
        ("V2_gutter_36pt", 720, False, "gutter=36pt (no mirror)"),
        ("V3_gutter_36pt_mirror", 720, True, "gutter=36pt + mirror (alternates side?)"),
    ]

    for label, gutter_tw, mirror, desc in variants:
        doc_xml = make_doc_xml(margin_l_tw, margin_r_tw, gutter_tw)
        settings_xml = make_settings_xml(mirror)
        try:
            p = make_docx(label, doc_xml, settings_xml)
        except Exception as e:
            out[label] = {"build_error": str(e)}
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        entry = {"label": label, "desc": desc, "gutter_tw": gutter_tw, "mirror": mirror, **r}
        out[label] = entry
        print(f"  {label}: {desc}")
        print(f"    page 1 first_x={entry.get('page1_first_x')}")
        print(f"    page 2 first_x={entry.get('page2_first_x')}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
