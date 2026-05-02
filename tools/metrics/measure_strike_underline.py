"""§4.x round 39 — strike / dstrike / underline width verify.

Quick test: do w:strike, w:dstrike, w:u affect text advance widths?
Expected: no effect (visual decoration only). Verify empirically.

Probe: "abcd" at fs=12 TNR with each decoration variant.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\strike_underline_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\strike_underline.json")
os.makedirs(OUT_DIR, exist_ok=True)

LATIN_FONT = "Times New Roman"
CJK_FONT = "ＭＳ 明朝"


def make_run(text, sz_val, strike=False, dstrike=False, u_val=None, b=False, i=False):
    extras = ""
    if strike:
        extras += "<w:strike/>"
    if dstrike:
        extras += "<w:dstrike/>"
    if u_val is not None:
        extras += f'<w:u w:val="{u_val}"/>'
    if b:
        extras += "<w:b/>"
    if i:
        extras += "<w:i/>"
    return ('<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'{extras}'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r>')


def make_doc_xml(runs_xml, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            f'{runs_xml}</w:p>'
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
        if not xs: return {"error": "no chars"}
        char_advances = []
        for i in range(len(xs) - 1):
            char_advances.append({"ch": xs[i][0], "adv": round(xs[i+1][1] - xs[i][1], 3)})
        return {"n_chars": len(xs), "advs": char_advances,
                "first_x": xs[0][1], "last_x": xs[-1][1],
                "span": round(xs[-1][1] - xs[0][1], 3)}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    page_w_tw = int((400 + 170) * 20)
    margin_tw = 170 * 10

    variants = [
        ("V1_plain",    make_run("abcd", sz_val), "no decoration"),
        ("V2_strike",   make_run("abcd", sz_val, strike=True), "w:strike"),
        ("V3_dstrike",  make_run("abcd", sz_val, dstrike=True), "w:dstrike (double strike)"),
        ("V4_u_single", make_run("abcd", sz_val, u_val="single"), "w:u single"),
        ("V5_u_double", make_run("abcd", sz_val, u_val="double"), "w:u double"),
        ("V6_b",        make_run("abcd", sz_val, b=True), "w:b (bold) — different glyph metrics"),
        ("V7_i",        make_run("abcd", sz_val, i=True), "w:i (italic) — different glyph metrics"),
    ]

    print(f"Test: decoration / formatting effect on advance widths\n")
    for label, run_xml, desc in variants:
        doc_xml = make_doc_xml(run_xml, page_w_tw, margin_tw)
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
        advs = entry.get("advs", [])
        adv_str = " ".join(f"{c['ch']!r}={c['adv']:.1f}" for c in advs)
        span = entry.get("span")
        print(f"  {label}: {desc}")
        print(f"    advs: {adv_str}, span={span}pt")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
