"""§5.X round 63 — w:cols multi-column page layout.

ECMA-376 §17.6.4 — section `<w:cols>` element. Default 1 col.
Multi-col with `<w:cols w:num="2" w:space="..."/>` divides content
area into N columns separated by space.

Probes:
  V1 1 col (default): reference, content fills full width
  V2 2 col, default space: text flows col 1 → col 2 with default gap
  V3 2 col, space=1440tw (= 36pt gap)
  V4 2 col, equalWidth=0, custom widths

Measure: x positions of chars to detect column boundaries.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\cols_layout_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\cols_layout.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(probe, sz_val, page_w_tw, margin_tw, cols_xml):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{probe}</w:t></w:r></w:p>'
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
        time.sleep(0.2)
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
        # Group by y to identify lines (might also detect columns by x jumps)
        ys = sorted(set(round(x[2], 1) for x in xs))
        x_min = min(x[1] for x in xs)
        x_max = max(x[1] for x in xs)
        return {"n_chars": len(xs),
                "first_x": xs[0][1], "first_y": xs[0][2],
                "x_range": [x_min, x_max],
                "n_unique_y": len(ys),
                "y_range": [min(ys), max(ys)],
                "first_8": [{"ch": xs[i][0], "x": xs[i][1], "y": xs[i][2]} for i in range(min(8, len(xs)))],
                "last_3": [{"ch": xs[i][0], "x": xs[i][1], "y": xs[i][2]} for i in range(max(0, len(xs)-3), len(xs))]}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    margin_tw = 170 * 10
    page_w_tw = int((400 + 170) * 20)  # 8500tw page, content = 400pt

    probe = "漢" * 60  # 60 chars to fill multiple lines

    print(f"Multi-column test, fs=12, content_w=400pt, probe=漢×60\n")

    variants = [
        ("V1_1col",
         '<w:cols w:space="720"/>',
         "1 col default"),
        ("V2_2col",
         '<w:cols w:num="2" w:space="720"/>',
         "2 col, space=720tw=36pt"),
        ("V3_2col_wide",
         '<w:cols w:num="2" w:space="1440"/>',
         "2 col, space=1440tw=72pt"),
        ("V4_3col",
         '<w:cols w:num="3" w:space="720"/>',
         "3 col"),
    ]

    for label, cols_xml, desc in variants:
        doc_xml = make_doc_xml(probe, sz_val, page_w_tw, margin_tw, cols_xml)
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
        print(f"  {label}: {desc}")
        print(f"    n_chars={entry.get('n_chars')} x_range={entry.get('x_range')} y_range={entry.get('y_range')}")
        print(f"    first 4: {[{'ch': c['ch'], 'x': c['x'], 'y': c['y']} for c in entry.get('first_8', [])[:4]]}")
        print(f"    last 3: {entry.get('last_3', [])}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
