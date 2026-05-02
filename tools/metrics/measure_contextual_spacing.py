"""§5.x round 36 — contextualSpacing paragraph property.

ECMA-376 §17.3.1.7 — `<w:contextualSpacing>` suppresses spacing
between consecutive paragraphs of the SAME pStyle. Common in list
styles to make item lists tight.

Test: 3 consecutive paragraphs with sa=200tw (10pt), sb=200tw (10pt)
each. Compare paragraph-to-paragraph distance with/without contextualSpacing.

Variants:
  V1 default: no contextualSpacing → standard collapse rule (max(sa_i, sb_{i+1}))
  V2 ON same style: all 3 paras same pStyle "MyStyle" with contextualSpacing=1
  V3 ON different style: P1 = "MyStyle1", P2,P3 = "MyStyle2", contextualSpacing on
      MyStyle2. Suppression should NOT fire between P1 and P2 (different styles)
      but SHOULD fire between P2 and P3 (same MyStyle2).
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\contextual_spacing_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\contextual_spacing.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_paragraph(text, sz_val, p_style_id="", ppr_extra=""):
    style_ref = f'<w:pStyle w:val="{p_style_id}"/>' if p_style_id else ''
    return ('<w:p>'
            f'<w:pPr>{style_ref}{ppr_extra}'
            '<w:spacing w:before="200" w:after="200" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>')


def make_doc_xml(paras_xml, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{paras_xml}'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml(sz_val, custom_styles=""):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
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
            # Get Y position of each paragraph
            n_paras = d.Paragraphs.Count
            ys = []
            for pi in range(1, n_paras + 1):
                p = d.Paragraphs(pi)
                y = float(p.Range.Information(6))
                ys.append(y)
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        gaps = [round(ys[i+1] - ys[i], 3) for i in range(len(ys) - 1)]
        return {"n_paras": len(ys), "para_ys": ys, "gaps": gaps}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    fs_pt = 12.0
    margin_tw = 170 * 10
    page_w_tw = int((400 + 170) * 20)

    # V1 default — 3 paras, no contextualSpacing
    paras_v1 = "".join([
        make_paragraph("段落1", sz_val),
        make_paragraph("段落2", sz_val),
        make_paragraph("段落3", sz_val),
    ])
    d1 = make_doc_xml(paras_v1, page_w_tw, margin_tw)
    s1 = make_styles_xml(sz_val)

    # V2 — all 3 paras same MyStyle with contextualSpacing
    custom_v2 = ('<w:style w:type="paragraph" w:styleId="MyStyle">'
                 '<w:name w:val="MyStyle"/>'
                 '<w:pPr><w:contextualSpacing/>'
                 '<w:spacing w:before="200" w:after="200"/>'
                 '</w:pPr></w:style>')
    paras_v2 = "".join([
        make_paragraph("段落1", sz_val, p_style_id="MyStyle"),
        make_paragraph("段落2", sz_val, p_style_id="MyStyle"),
        make_paragraph("段落3", sz_val, p_style_id="MyStyle"),
    ])
    d2 = make_doc_xml(paras_v2, page_w_tw, margin_tw)
    s2 = make_styles_xml(sz_val, custom_styles=custom_v2)

    # V3 — P1 different, P2 P3 same MyStyle with contextualSpacing
    custom_v3 = ('<w:style w:type="paragraph" w:styleId="StyleA">'
                 '<w:name w:val="StyleA"/>'
                 '<w:pPr><w:spacing w:before="200" w:after="200"/></w:pPr></w:style>'
                 '<w:style w:type="paragraph" w:styleId="StyleB">'
                 '<w:name w:val="StyleB"/>'
                 '<w:pPr><w:contextualSpacing/>'
                 '<w:spacing w:before="200" w:after="200"/>'
                 '</w:pPr></w:style>')
    paras_v3 = "".join([
        make_paragraph("段落1", sz_val, p_style_id="StyleA"),
        make_paragraph("段落2", sz_val, p_style_id="StyleB"),
        make_paragraph("段落3", sz_val, p_style_id="StyleB"),
    ])
    d3 = make_doc_xml(paras_v3, page_w_tw, margin_tw)
    s3 = make_styles_xml(sz_val, custom_styles=custom_v3)

    # V4 — pPr-level contextualSpacing on P2 only
    paras_v4 = "".join([
        make_paragraph("段落1", sz_val),
        make_paragraph("段落2", sz_val, ppr_extra='<w:contextualSpacing/>'),
        make_paragraph("段落3", sz_val),
    ])
    d4 = make_doc_xml(paras_v4, page_w_tw, margin_tw)
    s4 = make_styles_xml(sz_val)

    variants = [
        ("V1_default",       d1, s1, "no contextualSpacing"),
        ("V2_all_same_style", d2, s2, "all 3 paras same MyStyle + contextualSpacing"),
        ("V3_different_styles", d3, s3, "P1 StyleA, P2-P3 StyleB(ctxSp); only P2-P3 should suppress"),
        ("V4_pPr_only",      d4, s4, "P2 has pPr-level contextualSpacing"),
    ]

    print(f"Test: 3 paragraphs each with sa=200tw (10pt), sb=200tw (10pt) = total 200tw collapse\n")
    for label, doc_xml, styles_xml, desc in variants:
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
        entry = {"label": label, "desc": desc, **r}
        out[label] = entry
        ys = entry.get("para_ys", [])
        gaps = entry.get("gaps", [])
        print(f"  {label}: {desc}")
        print(f"    para_ys = {ys}")
        print(f"    gaps = {gaps}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
