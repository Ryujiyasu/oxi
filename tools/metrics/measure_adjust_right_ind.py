"""§4.x round 44 — w:adjustRightInd grid adjustment.

ECMA-376 §17.3.1.1 — `<w:adjustRightInd>` paragraph property.
"specifies whether the right indent shall be automatically adjusted
when text within the paragraph is followed by a [grid]".

Test: paragraph with explicit right indent + docGrid linePitch.
Vary adjustRightInd flag and observe line content end position.

Probe: 漢×N filling paragraph with right indent of 30pt.
docGrid linePitch=300 (15pt grid).

Variants:
  V1 default
  V2 val=1 (ON explicit)
  V3 val=0 (OFF)
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\adjust_right_ind_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\adjust_right_ind.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(probe, sz_val, page_w_tw, margin_tw, right_ind_tw, adjust=None):
    a_xml = ""
    if adjust is True:
        a_xml = '<w:adjustRightInd w:val="1"/>'
    elif adjust is False:
        a_xml = '<w:adjustRightInd w:val="0"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr>{a_xml}<w:ind w:right="{right_ind_tw}"/>'
            '<w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{probe}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360" w:type="lines"/>'
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
                    xs.append((t, float(c.Information(5)), float(c.Information(6))))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        ys = sorted(set(x[2] for x in xs))
        lines = []
        for y in ys:
            line_chars = sorted([(t, x) for t, x, yy in xs if abs(yy - y) < 0.5], key=lambda v: v[1])
            lines.append({
                "y": y, "n": len(line_chars),
                "first_x": line_chars[0][1] if line_chars else None,
                "last_x": line_chars[-1][1] if line_chars else None,
            })
        return {"n_total": len(xs), "n_lines": len(lines), "lines": lines}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    fs_pt = 12.0
    margin_tw = 170 * 10  # 85pt margins
    right_ind_pt = 30.0   # 30pt right indent
    right_ind_tw = int(right_ind_pt * 20)  # 600 tw

    cw_pt = 300.0  # generous content width
    page_w_tw = int((cw_pt + 170) * 20)
    # Effective content (minus right indent) = 300 - 30 = 270pt
    # Filled with 漢 chars: 270/12 = 22 chars max

    probe = "漢" * 30  # 30 chars to force wrap

    print(f"Probe: 漢×30 (natural=360pt) with right_ind=30pt, cw=300pt")
    print(f"  Effective line width = 300 - 30 = 270pt → 22 chars × 12pt fits")
    print(f"  Right edge with ind = 85 + 270 = 355pt (vs no ind: 85+300=385)")
    print(f"  docGrid linePitch=360 (line-based grid)\n")

    variants = [
        ("V1_default", None,  "no setting"),
        ("V2_on",      True,  "<w:adjustRightInd w:val=\"1\"/>"),
        ("V3_off",     False, "<w:adjustRightInd w:val=\"0\"/>"),
    ]

    for label, flag, desc in variants:
        doc_xml = make_doc_xml(probe, sz_val, page_w_tw, margin_tw, right_ind_tw, flag)
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
        entry = {"label": label, "desc": desc, "flag": flag, **r}
        out[label] = entry
        lines = entry.get("lines", [])
        print(f"  {label}: {desc}")
        for i, l in enumerate(lines):
            print(f"    line {i+1}: n={l['n']} first_x={l['first_x']} last_x={l['last_x']}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
