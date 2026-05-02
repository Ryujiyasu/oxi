"""§5.X round 73 — firstLine vs hanging indent.

ECMA-376 §17.3.1.12 — `<w:ind>` paragraph indent attributes:
  w:left      — left margin offset for paragraph
  w:right     — right margin offset
  w:firstLine — first line additional indent (positive = first line indented further right)
  w:hanging   — hanging indent (positive = first line shifted LEFT relative to rest)

firstLine and hanging are mutually exclusive. hanging = -firstLine (effectively).

Probes (multi-line content):
  V1 plain (no ind): all lines start at left margin
  V2 firstLine=400tw (20pt): line 1 starts +20, lines 2+ at left
  V3 hanging=400tw (20pt): line 1 at left, lines 2+ at +20
  V4 left=400 + firstLine=200: line 1 = 400+200=600, lines 2+ = 400
  V5 left=400 + hanging=200: line 1 = 400-200=200, lines 2+ = 400
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\first_line_ind_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\first_line_ind.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(ind_attrs, page_w_tw, margin_tw):
    """Long paragraph forcing line wrap."""
    ind_str = " ".join(f'{k}="{v}"' for k, v in ind_attrs.items()) if ind_attrs else ""
    ind_xml = f'<w:ind {ind_str}/>' if ind_attrs else ''
    long_text = "あいうえおかきくけこさしすせそたちつてと"  # 20 chars
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr>{ind_xml}<w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            '<w:sz w:val="24"/></w:rPr>'
            f'<w:t>{long_text}</w:t></w:r></w:p>'
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
        ys = sorted(set(round(x[2], 1) for x in xs))
        lines = []
        for y in ys:
            line_chars = sorted([(t, x) for t, x, yy in xs if abs(yy - y) < 0.5], key=lambda v: v[1])
            lines.append({"y": y, "n": len(line_chars),
                          "first_x": line_chars[0][1] if line_chars else None,
                          "last_x": line_chars[-1][1] if line_chars else None,
                          "first_chars": "".join(t for t, _ in line_chars[:6])})
        return {"n_total": len(xs), "n_lines": len(lines), "lines": lines}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    margin_tw = 170 * 10  # 85pt left margin
    page_w_tw = int((200 + 170) * 20)  # narrow content (200pt) to force wrap

    print(f"firstLine vs hanging indent test, fs=12 MS Mincho, content_w=200pt\n")
    print(f"  Probe: 20 chars × 12pt = 240pt (forces wrap)\n")

    variants = [
        ("V1_plain", None, "no indent"),
        ("V2_firstLine_400", {"w:firstLine": "400"}, "firstLine=400tw=20pt"),
        ("V3_hanging_400", {"w:hanging": "400"}, "hanging=400tw=20pt"),
        ("V4_left_400_firstLine_200",
         {"w:left": "400", "w:firstLine": "200"},
         "left=400 + firstLine=200 (line1=600, line2+=400)"),
        ("V5_left_400_hanging_200",
         {"w:left": "400", "w:hanging": "200"},
         "left=400 + hanging=200 (line1=200, line2+=400)"),
    ]

    for label, ind, desc in variants:
        doc_xml = make_doc_xml(ind, page_w_tw, margin_tw)
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
        lines = entry.get("lines", [])
        print(f"  {label}: {desc}")
        for i, l in enumerate(lines):
            print(f"    line {i+1}: y={l['y']} first_x={l['first_x']} chars={l['first_chars']!r}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
