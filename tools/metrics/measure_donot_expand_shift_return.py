"""§4.x round 43 — doNotExpandShiftReturn justification behavior.

ECMA-376 §17.15.1.40 — settings.xml `<w:doNotExpandShiftReturn>`
controls whether a line ending in Shift+Enter (soft line break,
`<w:br/>`) is justified to full width when paragraph alignment
is jc=both.

Default per ECMA: ON (lines ending in soft break NOT justified).

Test: probe with two lines separated by <w:br/>, jc=both.
  Line 1: 漢×10 (= 120pt natural)
  <w:br/>
  Line 2: 漢×5 (= 60pt natural)

cw = 200pt (much wider than both lines' natural).
- Default (doNotExpandShiftReturn ON): line 1 should NOT justify, last 漢 at x=85+120=205pt
- val=0 (OFF): line 1 SHOULD justify to full width, last 漢 at x=85+200=285pt

Measure last_x of line 1.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\donot_expand_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\donot_expand.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(sz_val, page_w_tw, margin_tw):
    """Paragraph with two runs separated by <w:br/>."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="both"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            '<w:t>漢漢漢漢漢漢漢漢漢漢</w:t>'
            '<w:br/>'
            '<w:t>漢漢漢漢漢</w:t>'
            '</w:r></w:p>'
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
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml(donot_expand=None):
    """donot_expand: None = no element, True = present (val=1), False = val=0."""
    inner = ""
    if donot_expand is True:
        inner = '<w:doNotExpandShiftReturn/>'
    elif donot_expand is False:
        inner = '<w:doNotExpandShiftReturn w:val="0"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'{inner}'
            '</w:settings>')


def make_docx(label, doc_xml, styles_xml, settings_xml):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
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
            advs = []
            for i in range(len(line_chars) - 1):
                advs.append(round(line_chars[i+1][1] - line_chars[i][1], 3))
            lines.append({
                "y": y, "n": len(line_chars),
                "first_x": line_chars[0][1] if line_chars else None,
                "last_x": line_chars[-1][1] if line_chars else None,
                "avg_adv": round(sum(advs)/len(advs), 2) if advs else None,
            })
        return {"n_total": len(xs), "n_lines": len(lines), "lines": lines}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    fs_pt = 12.0
    margin_tw = 170 * 10
    cw_pt = 200.0  # wider than both lines' natural
    page_w_tw = int((cw_pt + 170) * 20)

    print(f"Probe: 漢×10 + <w:br/> + 漢×5 at fs=12, jc=both, cw=200pt")
    print(f"  Line 1 natural = 120pt, Line 2 natural = 60pt")
    print(f"  Right edge = 85 + 200 = 285pt")
    print(f"  If line 1 justified: last 漢 at x ≈ 285pt, advance ≈ 22pt (= 200/9)")
    print(f"  If line 1 NOT justified: last 漢 at x ≈ 85+9*12 = 193pt\n")

    variants = [
        ("V1_default", None,  "no setting"),
        ("V2_on",      True,  "<w:doNotExpandShiftReturn/> (val=1)"),
        ("V3_off",     False, "<w:doNotExpandShiftReturn w:val=\"0\"/>"),
    ]

    for label, flag, desc in variants:
        doc_xml = make_doc_xml(sz_val, page_w_tw, margin_tw)
        styles_xml = make_styles_xml(sz_val)
        settings_xml = make_settings_xml(flag)
        try:
            p = make_docx(label, doc_xml, styles_xml, settings_xml)
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
            justified = "JUSTIFIED" if l.get("last_x") and l["last_x"] > 270 else "natural"
            print(f"    line {i+1}: n={l['n']} first_x={l['first_x']} last_x={l['last_x']} avg_adv={l['avg_adv']} ({justified})")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
