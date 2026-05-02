"""§5.CC round 68 — w:br w:type="column" column break.

ECMA-376 §17.3.3.1 — `<w:br w:type="column"/>` forces advance to
next column in multi-column layout (without filling current vertically).

Probes (2-col layout, space=36pt):
  V1 plain "Col1Text" + "Col2Text" (no break, all in col 1)
  V2 "Col1Text" + col_break + "Col2Text"
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\column_break_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\column_break.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_run_with_break(text_before, text_after, sz_val, br_type=None):
    type_attr = f' w:type="{br_type}"' if br_type else ''
    return ('<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{text_before}</w:t>'
            f'<w:br{type_attr}/>'
            f'<w:t>{text_after}</w:t></w:r>')


def make_run_plain(text, sz_val):
    return ('<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r>')


def make_doc_xml(content_xml, page_w_tw, margin_tw, cols_xml):
    body = ('<w:p><w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'</w:pPr>{content_xml}</w:p>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            f'{cols_xml}'
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
        char_info = [{"ch": xs[i][0], "x": xs[i][1], "y": xs[i][2]} for i in range(len(xs))]
        return {"n_chars": len(xs), "chars": char_info}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    margin_tw = 170 * 10
    page_w_tw = int((400 + 170) * 20)
    cols_2 = '<w:cols w:num="2" w:space="720"/>'

    # Col 1 spans 85-267 (col_w=182), col 2 spans 303-485

    # V1: plain "Col1Text" + "Col2Text" — both in col 1
    v1 = make_run_plain("Col1Text") if False else (
         '<w:r><w:rPr>'
         f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
         f'<w:sz w:val="{sz_val}"/></w:rPr>'
         f'<w:t>Col1TextCol2Text</w:t></w:r>')

    # V2: with col break: "Col1Text" + col_break + "Col2Text"
    v2 = make_run_with_break("Col1Text", "Col2Text", sz_val, br_type="column")

    # V3: with page break (control)
    v3 = make_run_with_break("Col1Text", "Col2Text", sz_val, br_type="page")

    print(f"Column break test, fs=12 MS Mincho, 2-col layout\n")

    variants = [
        ("V1_no_break", v1, "no break (both in col 1)"),
        ("V2_col_break", v2, "Col1Text + col break + Col2Text"),
        ("V3_page_break", v3, "Col1Text + page break + Col2Text (control)"),
    ]

    for label, body_xml, desc in variants:
        doc_xml = make_doc_xml(body_xml, page_w_tw, margin_tw, cols_2)
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
        chars = entry.get("chars", [])
        print(f"  {label}: {desc}")
        for c in chars[:18]:
            print(f"    {c['ch']!r}: x={c['x']} y={c['y']}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
