"""§4.x round 33 — w:vanish (hidden text) layout effect.

ECMA-376 §17.3.2.43 — `<w:vanish/>` run property hides text from
display. Question: does hidden text occupy layout space (advance),
affect line wrap, or use zero width?

Probes:
  V1: 漢×5 (baseline visible)
  V2: 漢×3 + <vanish>漢×3</vanish> + 漢×2 (hidden middle, mixed)
  V3: 漢×3 + <vanish>漢×10</vanish> + 漢×2 (large hidden block)
  V4: <vanish>全文</vanish> (entire para hidden)

For V2/V3, measure visible char positions to see if hidden chars
shifted them or were skipped (zero width).
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\vanish_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\vanish.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(runs_xml, sz_val, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            f'{runs_xml}'
            '</w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_run(text, sz_val, vanish=False):
    """Make a run with optional w:vanish."""
    vanish_xml = '<w:vanish/>' if vanish else ''
    return ('<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'{vanish_xml}'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r>')


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
        # Show all chars and their positions
        char_info = [{"ch": t, "x": x, "y": y} for t, x, y in xs]
        return {"n_chars": len(xs), "chars": char_info[:30]}  # cap at 30 for readability
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    fs_pt = 12.0
    margin_tw = 170 * 10
    page_w_tw = int((400 + 170) * 20)

    # V1: 漢×5 baseline
    v1_runs = make_run("漢漢漢漢漢", sz_val)
    v1_doc = make_doc_xml(v1_runs, sz_val, page_w_tw, margin_tw)

    # V2: 漢×3 visible + 漢×3 hidden + 漢×2 visible
    v2_runs = (make_run("漢漢漢", sz_val, vanish=False)
               + make_run("漢漢漢", sz_val, vanish=True)
               + make_run("漢漢", sz_val, vanish=False))
    v2_doc = make_doc_xml(v2_runs, sz_val, page_w_tw, margin_tw)

    # V3: 漢×3 visible + 漢×10 hidden + 漢×2 visible
    v3_runs = (make_run("漢漢漢", sz_val, vanish=False)
               + make_run("漢漢漢漢漢漢漢漢漢漢", sz_val, vanish=True)
               + make_run("漢漢", sz_val, vanish=False))
    v3_doc = make_doc_xml(v3_runs, sz_val, page_w_tw, margin_tw)

    # V4: All hidden
    v4_runs = make_run("漢漢漢漢漢", sz_val, vanish=True)
    v4_doc = make_doc_xml(v4_runs, sz_val, page_w_tw, margin_tw)

    # V5: short visible + 200-char hidden + short visible (force test wrap)
    v5_runs = (make_run("漢漢漢", sz_val, vanish=False)
               + make_run("漢" * 200, sz_val, vanish=True)
               + make_run("漢漢漢", sz_val, vanish=False))
    v5_doc = make_doc_xml(v5_runs, sz_val, page_w_tw, margin_tw)

    variants = [
        ("V1_baseline", v1_doc, "5 visible (baseline)"),
        ("V2_mid_hidden", v2_doc, "3v + 3h + 2v"),
        ("V3_large_hidden", v3_doc, "3v + 10h + 2v"),
        ("V4_all_hidden", v4_doc, "all hidden"),
        ("V5_force_wrap", v5_doc, "3v + 200h + 3v (test wrap)"),
    ]

    print(f"Test: w:vanish hidden text layout effect at fs=12 MS Mincho\n")
    for label, doc_xml, desc in variants:
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
        n = entry.get("n_chars", "?")
        chars = entry.get("chars", [])
        print(f"\n{label} ({desc}): n_chars={n}")
        # Print first 8 char positions
        for c in chars[:10]:
            print(f"    ch={c['ch']!r} x={c['x']} y={c['y']}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
