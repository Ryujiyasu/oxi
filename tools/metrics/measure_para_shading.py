"""§5.X round 71 — w:shd paragraph shading.

ECMA-376 §17.3.5.31 — `<w:shd w:val="..." w:fill="..." w:color="..."/>`
in pPr applies background fill to paragraph.

Probes:
  V1 plain (no shd): reference
  V2 shd val=clear fill=FFFF00 (yellow background)
  V3 shd val=pct50 fill=000000 (50% gray pattern)
  V4 shd combined with border
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\para_shading_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\para_shading.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(ppr_extra, page_w_tw, margin_tw):
    p1 = ('<w:p><w:pPr><w:jc w:val="left"/>'
          f'{ppr_extra}'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
          '</w:pPr>'
          '<w:r><w:rPr>'
          f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
          '<w:sz w:val="24"/></w:rPr>'
          '<w:t>あいうえお</w:t></w:r></w:p>')
    p2 = ('<w:p><w:pPr><w:jc w:val="left"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
          '</w:pPr>'
          '<w:r><w:rPr>'
          f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
          '<w:sz w:val="24"/></w:rPr>'
          '<w:t>second</w:t></w:r></w:p>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{p1}{p2}'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
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
                    xs.append((t, float(c.Information(5)), float(c.Information(6))))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        return {"n_chars": len(xs), "first_x": xs[0][1], "first_y": xs[0][2],
                "last_3": [{"ch": xs[i][0], "x": xs[i][1], "y": xs[i][2]} for i in range(max(0, len(xs)-3), len(xs))]}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    margin_tw = 170 * 10
    page_w_tw = int((400 + 170) * 20)

    print(f"Paragraph shading test, fs=12 MS Mincho\n")

    variants = [
        ("V1_no_shd", "", "no shading (control)"),
        ("V2_yellow_fill",
         '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>',
         "shd clear fill=FFFF00 (yellow background)"),
        ("V3_gray_pattern",
         '<w:shd w:val="pct50" w:color="000000" w:fill="auto"/>',
         "shd pct50 (50% gray pattern)"),
        ("V4_shd_with_border",
         '<w:pBdr><w:top w:val="single" w:sz="12" w:space="4" w:color="000000"/>'
         '<w:bottom w:val="single" w:sz="12" w:space="4" w:color="000000"/></w:pBdr>'
         '<w:shd w:val="clear" w:color="auto" w:fill="CCCCCC"/>',
         "shd + top/bottom borders"),
    ]

    for label, ppr_extra, desc in variants:
        doc_xml = make_doc_xml(ppr_extra, page_w_tw, margin_tw)
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
        print(f"    n={entry.get('n_chars')} first_x={entry.get('first_x')} first_y={entry.get('first_y')}")
        for c in entry.get("last_3", []):
            print(f"      last: {c['ch']!r} x={c['x']} y={c['y']}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
