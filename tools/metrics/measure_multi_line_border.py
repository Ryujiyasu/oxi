"""§5.EE round 75 — multi-line paragraph border space.

Round 70 caveat: with multi-line paragraphs, is top/bottom border
space applied once per paragraph or per line?

Probes (cw narrow forces wrap):
  V1 plain 1-line paragraph + plain (control)
  V2 1-line paragraph with top+bottom border + plain
  V3 multi-line paragraph with top+bottom border + plain
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\multi_line_border_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\multi_line_border.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_para(text, sz_val, with_border=False):
    pbdr = ('<w:pBdr>'
            '<w:top w:val="single" w:sz="12" w:space="4" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="12" w:space="4" w:color="000000"/>'
            '</w:pBdr>') if with_border else ''
    return ('<w:p><w:pPr><w:jc w:val="left"/>'
            f'{pbdr}'
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
            n_paras = d.Paragraphs.Count
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        ys = sorted(set(round(x[2], 1) for x in xs))
        lines = []
        for y in ys:
            line_chars = sorted([(t, x) for t, x, yy in xs if abs(yy - y) < 0.5], key=lambda v: v[1])
            lines.append({"y": y, "n": len(line_chars),
                          "first_chars": "".join(t for t, _ in line_chars[:5])})
        return {"n_total": len(xs), "n_paras": n_paras, "n_lines": len(lines),
                "lines": lines}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    margin_tw = 170 * 10
    page_w_tw = int((150 + 170) * 20)  # narrow 150pt content forces wrap

    short_text = "あ"  # 1 char = 1 line
    long_text = "あいうえおかきくけこさしすせそたちつてと"  # 20 chars, wraps to ~2-3 lines at 150pt

    # V1: short P1 + short P2 (no border)
    v1 = make_para(short_text, sz_val) + make_para(short_text, sz_val)

    # V2: short P1 with border + short P2
    v2 = make_para(short_text, sz_val, with_border=True) + make_para(short_text, sz_val)

    # V3: long P1 with border (multi-line) + short P2
    v3 = make_para(long_text, sz_val, with_border=True) + make_para(short_text, sz_val)

    # V4: long P1 (no border) + short P2 — control for compare
    v4 = make_para(long_text, sz_val) + make_para(short_text, sz_val)

    print(f"Multi-line border test, fs=12 MS Mincho, content_w=150pt\n")

    variants = [
        ("V1_short_no_border", make_doc_xml(v1, page_w_tw, margin_tw),
         "short P1 + short P2 (no border)"),
        ("V2_short_with_border", make_doc_xml(v2, page_w_tw, margin_tw),
         "short P1 with border + short P2"),
        ("V3_long_with_border", make_doc_xml(v3, page_w_tw, margin_tw),
         "long P1 (multi-line) with border + short P2"),
        ("V4_long_no_border", make_doc_xml(v4, page_w_tw, margin_tw),
         "long P1 (multi-line) no border + short P2"),
    ]

    for label, doc_xml, desc in variants:
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
        print(f"    n_lines={entry.get('n_lines')}")
        for i, l in enumerate(entry.get("lines", [])):
            print(f"      line {i+1}: y={l['y']} n={l['n']} chars={l['first_chars']!r}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
