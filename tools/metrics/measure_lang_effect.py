"""§4.6.2 round 55 — w:lang effect on autoSpaceDE.

ECMA-376 §17.3.2.20 — `<w:lang>` run property specifies language
for spell-check etc. Does it gate autoSpaceDE boundary?

Probes (fs=12 MS Mincho, probe 漢a):
  V1 default (no lang)
  V2 lang val="ja-JP"
  V3 lang val="en-US"
  V4 lang eastAsia="ja-JP"
  V5 lang eastAsia="en-US"
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\lang_effect_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\lang_effect.json")
os.makedirs(OUT_DIR, exist_ok=True)

LATIN_FONT = "Times New Roman"
CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(probe, sz_val, page_w_tw, margin_tw, lang_attrs):
    lang_xml = ""
    if lang_attrs:
        attrs = " ".join(f'w:{k}="{v}"' for k, v in lang_attrs.items())
        lang_xml = f'<w:lang {attrs}/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'{lang_xml}'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t xml:space="preserve">{probe}</w:t></w:r></w:p>'
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
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '</w:settings>')


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
        time.sleep(0.2)
        try:
            chars = d.Range().Characters
            xs = []
            for ci in range(1, chars.Count + 1):
                try:
                    c = chars(ci)
                    t = c.Text
                    if not t or any(ord(ch) < 32 for ch in t): continue
                    xs.append((t, float(c.Information(5))))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs or len(xs) < 2: return {"error": "no chars"}
        first_adv = round(xs[1][1] - xs[0][1], 3)
        return {"first_char": xs[0][0], "first_char_adv": first_adv}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    page_w_tw = int((400 + 170) * 20)
    margin_tw = 170 * 10
    probe = "漢a"

    print(f"Probe: 漢a (kanji + halfwidth a) at fs=12 TNR + MSM")
    print(f"Expected if boundary fires: 漢=15.0pt; if disabled: 漢=12.0pt\n")

    variants = [
        ("V1_no_lang",        None, "no lang attribute"),
        ("V2_lang_jaJP",      {"val": "ja-JP"}, "w:lang val=ja-JP"),
        ("V3_lang_enUS",      {"val": "en-US"}, "w:lang val=en-US"),
        ("V4_easia_jaJP",     {"eastAsia": "ja-JP"}, "w:lang eastAsia=ja-JP"),
        ("V5_easia_enUS",     {"eastAsia": "en-US"}, "w:lang eastAsia=en-US"),
        ("V6_combined_enUS",  {"val": "en-US", "eastAsia": "ja-JP"},
         "w:lang val=en-US eastAsia=ja-JP"),
    ]

    for label, lang_attrs, desc in variants:
        doc_xml = make_doc_xml(probe, sz_val, page_w_tw, margin_tw, lang_attrs)
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
        adv = r.get("first_char_adv")
        boundary_fires = (adv == 15.0)
        out[label] = {"label": label, "desc": desc, **r,
                      "boundary_fires": boundary_fires}
        print(f"  {label}: {desc}")
        print(f"    漢_adv={adv} {'(fires)' if boundary_fires else '(NO fire)' if adv == 12.0 else '(?)'}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
