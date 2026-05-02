"""§5.AA round 61 — proofErr / bookmark / smartTag transparency.

Common XML wrappers in baseline docs:
  <w:proofErr w:type="...">  — spell/grammar error markers
  <w:bookmarkStart w:id="..." w:name="..."/>
  <w:bookmarkEnd w:id="..."/>
  <w:smartTag w:uri="..." w:element="...">...runs...</w:smartTag>

All expected to be layout-transparent (metadata only).

Probes:
  V1 plain "Hello World" reference
  V2 with proofErr wrappers around runs
  V3 with bookmarkStart/End around runs
  V4 with smartTag wrapping runs
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\xml_wrappers_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\xml_wrappers.json")
os.makedirs(OUT_DIR, exist_ok=True)

LATIN_FONT = "Times New Roman"
CJK_FONT = "ＭＳ 明朝"


def make_run_text(text, sz_val):
    return ('<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r>')


def make_doc_xml(body_inner_xml, sz_val, page_w_tw, margin_tw):
    body = ('<w:p><w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'</w:pPr>{body_inner_xml}</w:p>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}'
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
                    xs.append((t, float(c.Information(5))))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        char_info = []
        for i in range(len(xs)):
            adv = round(xs[i+1][1] - xs[i][1], 3) if i < len(xs) - 1 else None
            char_info.append({"ch": repr(xs[i][0]), "x": xs[i][1], "adv": adv})
        return {"n_chars": len(xs), "chars": char_info}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    margin_tw = 170 * 10
    page_w_tw = int((400 + 170) * 20)

    print(f"XML wrapper transparency test, fs=12 TNR\n")

    # V1: plain run
    v1 = make_run_text("Hello World", sz_val)

    # V2: with proofErr around the run
    v2 = ('<w:proofErr w:type="spellStart"/>'
          + make_run_text("Hello World", sz_val)
          + '<w:proofErr w:type="spellEnd"/>')

    # V3: with bookmark wrap
    v3 = ('<w:bookmarkStart w:id="0" w:name="myBookmark"/>'
          + make_run_text("Hello World", sz_val)
          + '<w:bookmarkEnd w:id="0"/>')

    # V4: with smartTag wrapping the run
    v4 = ('<w:smartTag w:uri="urn:test" w:element="testTag">'
          + make_run_text("Hello World", sz_val)
          + '</w:smartTag>')

    variants = [
        ("V1_plain",    v1, "plain run (control)"),
        ("V2_proofErr", v2, "with proofErr wrapper"),
        ("V3_bookmark", v3, "with bookmarkStart/End"),
        ("V4_smartTag", v4, "with smartTag wrapper"),
    ]

    for label, body_xml, desc in variants:
        doc_xml = make_doc_xml(body_xml, sz_val, page_w_tw, margin_tw)
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
        n = entry.get("n_chars", "?")
        chars_str = " ".join(f"{c['ch']}={c['adv']}" for c in chars[:6])
        print(f"  {label}: {desc}")
        print(f"    n={n} | {chars_str}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
