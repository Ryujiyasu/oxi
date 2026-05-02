"""§5.EE round 74 — w:between border.

ECMA-376 §17.3.1.24 — `<w:between>` sub-element of pBdr renders
border between consecutive paragraphs sharing the SAME style.

Probes:
  V1 plain 3 paragraphs (no border)
  V2 3 paragraphs with between=1.5pt sz=12 space=4
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\between_border_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\between_border.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_para(text, p_style_id, sz_val, with_between=False):
    pbdr = ('<w:pBdr><w:between w:val="single" w:sz="12" w:space="4" w:color="000000"/></w:pBdr>'
            if with_between else '')
    style_ref = f'<w:pStyle w:val="{p_style_id}"/>' if p_style_id else ''
    return ('<w:p><w:pPr>'
            f'{style_ref}'
            f'{pbdr}'
            '<w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>')


def make_doc_xml(paras_inner, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{paras_inner}'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml(custom_styles=""):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            '<w:sz w:val="24"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults>'
            f'{custom_styles}'
            '</w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>')


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
                    text = rng.Text[:30] if rng.Text else ""
                    para_info.append({"i": pi, "y": y, "text_preview": text})
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        gaps = [round(para_info[i+1]["y"] - para_info[i]["y"], 3)
                for i in range(len(para_info)-1)]
        return {"n_paras": len(para_info), "paragraphs": para_info, "gaps": gaps}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24

    # MyStyle for between-paragraph border test
    custom_v2 = ('<w:style w:type="paragraph" w:styleId="MyStyle">'
                 '<w:name w:val="MyStyle"/>'
                 '<w:pPr/>'
                 '</w:style>')

    # V1: 3 plain paragraphs (no style, no border)
    paras_v1 = "".join([
        make_para("Para1", None, sz_val),
        make_para("Para2", None, sz_val),
        make_para("Para3", None, sz_val),
    ])
    s1 = make_styles_xml()

    # V2: 3 paragraphs same MyStyle with between border
    paras_v2 = "".join([
        make_para("Para1", "MyStyle", sz_val, with_between=True),
        make_para("Para2", "MyStyle", sz_val, with_between=True),
        make_para("Para3", "MyStyle", sz_val, with_between=True),
    ])
    s2 = make_styles_xml(custom_styles=custom_v2)

    # V3: 3 paragraphs DIFFERENT styles with between (between should NOT fire)
    custom_v3 = ('<w:style w:type="paragraph" w:styleId="StyleA">'
                 '<w:name w:val="StyleA"/><w:pPr/></w:style>'
                 '<w:style w:type="paragraph" w:styleId="StyleB">'
                 '<w:name w:val="StyleB"/><w:pPr/></w:style>')
    paras_v3 = "".join([
        make_para("Para1", "StyleA", sz_val, with_between=True),
        make_para("Para2", "StyleB", sz_val, with_between=True),
        make_para("Para3", "StyleA", sz_val, with_between=True),
    ])
    s3 = make_styles_xml(custom_styles=custom_v3)

    margin_tw = 170 * 10
    page_w_tw = int((400 + 170) * 20)

    print(f"Between border test, fs=12 MS Mincho\n")

    variants = [
        ("V1_no_border", make_doc_xml(paras_v1, page_w_tw, margin_tw), s1, "3 plain paragraphs"),
        ("V2_same_style_between", make_doc_xml(paras_v2, page_w_tw, margin_tw), s2,
         "3 paras same MyStyle + between border"),
        ("V3_mixed_styles_between", make_doc_xml(paras_v3, page_w_tw, margin_tw), s3,
         "Para1 StyleA, Para2 StyleB, Para3 StyleA + between (different styles)"),
    ]

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
        print(f"  {label}: {desc}")
        for p in entry.get("paragraphs", []):
            print(f"    para[{p['i']}]: y={p['y']} text={p['text_preview']!r}")
        print(f"    gaps={entry.get('gaps')}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
