"""§4.6.2.1 round 35 — autoSpaceDN inheritance chain.

Round 30 verified autoSpaceDE inheritance via 6 variants. Round 35
confirms autoSpaceDN follows the same chain via the same variants
but with DN flag and digit probe.

Probe: 漢1 (kana→digit), expected 漢=15.0 with DN=on, 12.0 with DN=off.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\autospace_dn_inherit_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\autospace_dn_inherit.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"
LATIN_FONT = "Times New Roman"


def make_doc_xml(probe, sz_val, page_w_tw, margin_tw, p_style_id="", ppr_extra=""):
    style_ref = f'<w:pStyle w:val="{p_style_id}"/>' if p_style_id else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr>{style_ref}{ppr_extra}<w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t xml:space="preserve">{probe}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml(sz_val, doc_default_dn_off=False, custom_styles=""):
    dd_dn = '<w:autoSpaceDN w:val="0"/>' if doc_default_dn_off else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            f'<w:pPrDefault><w:pPr>{dd_dn}</w:pPr></w:pPrDefault>'
            '</w:docDefaults>'
            f'{custom_styles}'
            '</w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '</w:settings>')


def make_docx(label, doc_xml, styles_xml):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
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
    probe = "漢1"

    # V1 — docDefaults DN=0
    s1 = make_styles_xml(sz_val, doc_default_dn_off=True)
    d1 = make_doc_xml(probe, sz_val, page_w_tw=11400, margin_tw=1700, p_style_id="")

    # V2 — Normal style DN=0
    custom_v2 = ('<w:style w:type="paragraph" w:styleId="Normal" w:default="1">'
                 '<w:name w:val="Normal"/>'
                 '<w:pPr><w:autoSpaceDN w:val="0"/></w:pPr>'
                 '</w:style>')
    s2 = make_styles_xml(sz_val, custom_styles=custom_v2)
    d2 = make_doc_xml(probe, sz_val, page_w_tw=11400, margin_tw=1700, p_style_id="Normal")

    # V3 — Normal DN=0 + Custom basedOn (no override)
    custom_v3 = ('<w:style w:type="paragraph" w:styleId="Normal" w:default="1">'
                 '<w:name w:val="Normal"/>'
                 '<w:pPr><w:autoSpaceDN w:val="0"/></w:pPr>'
                 '</w:style>'
                 '<w:style w:type="paragraph" w:styleId="Custom">'
                 '<w:name w:val="Custom"/>'
                 '<w:basedOn w:val="Normal"/>'
                 '<w:pPr></w:pPr>'
                 '</w:style>')
    s3 = make_styles_xml(sz_val, custom_styles=custom_v3)
    d3 = make_doc_xml(probe, sz_val, page_w_tw=11400, margin_tw=1700, p_style_id="Custom")

    # V4 — Normal DN=0, Custom DN=1 override
    custom_v4 = ('<w:style w:type="paragraph" w:styleId="Normal" w:default="1">'
                 '<w:name w:val="Normal"/>'
                 '<w:pPr><w:autoSpaceDN w:val="0"/></w:pPr>'
                 '</w:style>'
                 '<w:style w:type="paragraph" w:styleId="Custom">'
                 '<w:name w:val="Custom"/>'
                 '<w:basedOn w:val="Normal"/>'
                 '<w:pPr><w:autoSpaceDN w:val="1"/></w:pPr>'
                 '</w:style>')
    s4 = make_styles_xml(sz_val, custom_styles=custom_v4)
    d4 = make_doc_xml(probe, sz_val, page_w_tw=11400, margin_tw=1700, p_style_id="Custom")

    # V5 — Custom DN=0, pPr DN=1
    custom_v5 = ('<w:style w:type="paragraph" w:styleId="Normal" w:default="1">'
                 '<w:name w:val="Normal"/>'
                 '<w:pPr></w:pPr></w:style>'
                 '<w:style w:type="paragraph" w:styleId="Custom">'
                 '<w:name w:val="Custom"/>'
                 '<w:basedOn w:val="Normal"/>'
                 '<w:pPr><w:autoSpaceDN w:val="0"/></w:pPr>'
                 '</w:style>')
    s5 = make_styles_xml(sz_val, custom_styles=custom_v5)
    d5 = make_doc_xml(probe, sz_val, page_w_tw=11400, margin_tw=1700,
                      p_style_id="Custom",
                      ppr_extra='<w:autoSpaceDN w:val="1"/>')

    # V6 baseline default
    s6 = make_styles_xml(sz_val)
    d6 = make_doc_xml(probe, sz_val, page_w_tw=11400, margin_tw=1700)

    variants = [
        ("V6_baseline_default", d6, s6, 15.0, "default DN=on"),
        ("V1_docDef_DN0",       d1, s1, 12.0, "docDefaults DN=0"),
        ("V2_Normal_DN0",       d2, s2, 12.0, "Normal style DN=0 + use Normal"),
        ("V3_Custom_inherits",  d3, s3, 12.0, "Custom basedOn Normal(DN=0), no override"),
        ("V4_Custom_overrides", d4, s4, 15.0, "Custom basedOn Normal(DN=0), Custom DN=1"),
        ("V5_pPr_overrides",    d5, s5, 15.0, "Custom DN=0, pPr DN=1"),
    ]

    print(f"Probe: {probe!r} (kana→digit, expected 漢=15.0pt with DN=on, 12.0pt with DN=off)\n")
    for label, doc_xml, styles_xml, expected, desc in variants:
        try:
            p = make_docx(label, doc_xml, styles_xml)
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
        match = "✓" if adv == expected else ("✗" if adv is not None else "-")
        out[label] = {"label": label, "desc": desc, "expected": expected, "first_char_adv": adv, "match": match}
        print(f"  {label:>22} | {desc:>40} | expected={expected} observed={adv} {match}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
