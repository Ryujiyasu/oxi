"""§5.II round 86 — track changes (w:ins / w:del).

ECMA-376 §17.13.5.x — track changes elements:
  <w:ins w:id="..." w:author="..." w:date="...">  inserted run wrapper
  <w:del w:id="..." w:author="..." w:date="...">  deleted run wrapper
  <w:delText>...</w:delText> for deleted text

Probes:
  V1 plain "HelloWorld"
  V2 with w:ins wrapping "Inserted" between Hello/World
  V3 with w:del wrapping "Deleted" between Hello/World
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\track_changes_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\track_changes.json")
os.makedirs(OUT_DIR, exist_ok=True)

LATIN_FONT = "Times New Roman"
CJK_FONT = "ＭＳ 明朝"


def make_run(text, sz_val):
    return ('<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r>')


def make_ins_run(text, sz_val):
    return ('<w:ins w:id="0" w:author="Tester" w:date="2026-05-03T00:00:00Z">'
            f'{make_run(text, sz_val)}'
            '</w:ins>')


def make_del_run(text, sz_val):
    return ('<w:del w:id="0" w:author="Tester" w:date="2026-05-03T00:00:00Z">'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:delText>{text}</w:delText></w:r>'
            '</w:del>')


def make_doc_xml(content_xml, page_w_tw, margin_tw):
    body = ('<w:p><w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'</w:pPr>{content_xml}</w:p>')
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
        time.sleep(0.3)
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
            rev_count = 0
            try:
                rev_count = d.Revisions.Count
            except: pass
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        return {"n_chars": len(xs),
                "first_chars": "".join(x[0] for x in xs[:20]),
                "revisions_count": rev_count}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    margin_tw = 1700
    page_w_tw = int((400 + 170) * 20)

    print(f"Track changes test, fs=12 TNR\n")

    # V1 plain
    v1 = make_run("Hello", sz_val) + make_run("World", sz_val)

    # V2 ins inserted between
    v2 = (make_run("Hello", sz_val)
          + make_ins_run("Inserted", sz_val)
          + make_run("World", sz_val))

    # V3 del deleted between
    v3 = (make_run("Hello", sz_val)
          + make_del_run("Deleted", sz_val)
          + make_run("World", sz_val))

    variants = [
        ("V1_plain", v1, "plain HelloWorld"),
        ("V2_ins", v2, "Hello + ins(Inserted) + World"),
        ("V3_del", v3, "Hello + del(Deleted) + World"),
    ]

    for label, body_xml, desc in variants:
        doc_xml = make_doc_xml(body_xml, page_w_tw, margin_tw)
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
        print(f"    n_chars={entry.get('n_chars')} first_chars={entry.get('first_chars')!r}")
        print(f"    revisions={entry.get('revisions_count')}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
