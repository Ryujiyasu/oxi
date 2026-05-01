"""
§9 Footnote 17.5pt floor — pin the origin.

Prior result: footnote default lh = 17.5pt for MS Mincho 10.5pt (natural=13.617),
contradicting spec §9.1 "12pt Single" claim.

Hypothesis: Word applies a hidden "FootnoteText" style with implicit
`<w:spacing w:line="350" w:lineRule="atLeast"/>` (or similar floor mechanism).

Test: build docx with EXPLICIT styles.xml that overrides FootnoteText with
known lineSpacing values. If our explicit override flows through → floor
is just style inheritance. If 17.5pt persists despite override → floor is
hardcoded in Word's renderer.

Variants:
  V1: no styles.xml (baseline — should match 17.5pt prior data)
  V2: styles.xml with FootnoteText having spacing line=240 auto (= "Single")
  V3: styles.xml with FootnoteText having spacing line=200 auto (< natural)
  V4: styles.xml with FootnoteText having spacing line=200 exact
  V5: styles.xml with FootnoteText having NO spacing element
  V6: styles.xml with FootnoteText having spacing line=400 atLeast (= 20pt floor)
  V7: styles.xml with FootnoteText having spacing line=350 atLeast (try to match observed floor)

All variants use MS Mincho 10.5pt for body+footnote text, 3 footnotes per page.

Output: tools/metrics/output/footnote_floor_origin.json
"""
import json, os, sys, time, zipfile, uuid
from pathlib import Path
import pythoncom
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).with_name("output") / "footnote_floor_origin.json"
TMP_DIR = Path("pipeline_data") / "_footnote_floor_tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)


def content_types(with_styles):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
    ]
    if with_styles:
        parts.append('<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>')
    parts.append('</Types>')
    return "".join(parts)


PKG_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def doc_rels(with_styles):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
    ]
    if with_styles:
        parts.append('<Relationship Id="rSt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    parts.append('</Relationships>')
    return "".join(parts)


FONT = "ＭＳ 明朝"
SZ = 21


def rpr():
    return f'<w:rPr><w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/><w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/></w:rPr>'


def doc_xml():
    body_runs = ""
    for i in range(3):
        body_runs += (
            f'<w:r>{rpr()}<w:t xml:space="preserve">b{i+1}</w:t></w:r>'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{i+2}"/></w:r>'
        )
    body = f'<w:p><w:pPr>{rpr()}</w:pPr>{body_runs}</w:p>'
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}{sect}</w:body></w:document>')


def fn_xml(use_style):
    """Footnote text. If use_style, applies <w:pStyle val="FootnoteText"/>."""
    style_ref = '<w:pStyle w:val="FootnoteText"/>' if use_style else ""
    fn_entries = [
        '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for i in range(3):
        ppr = f'<w:pPr>{style_ref}{rpr()}</w:pPr>'
        fn_entries.append(
            f'<w:footnote w:id="{i+2}"><w:p>{ppr}'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            f'<w:r>{rpr()}<w:t xml:space="preserve"> b{i+1}</w:t></w:r></w:p></w:footnote>'
        )
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + "".join(fn_entries) +
            '</w:footnotes>')


def styles_xml(spacing_attrs):
    """spacing_attrs: e.g., 'w:line="240" w:lineRule="auto"' or None for empty pPr."""
    spacing_el = f'<w:spacing {spacing_attrs}/>' if spacing_attrs else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/>'
            '</w:style>'
            '<w:style w:type="paragraph" w:styleId="FootnoteText">'
            '<w:name w:val="footnote text"/>'
            '<w:basedOn w:val="Normal"/>'
            '<w:link w:val="FootnoteTextChar"/>'
            f'<w:pPr>{spacing_el}</w:pPr>'
            '</w:style>'
            '<w:style w:type="character" w:styleId="FootnoteReference">'
            '<w:name w:val="footnote reference"/>'
            '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
            '</w:style>'
            '<w:style w:type="character" w:styleId="FootnoteTextChar">'
            '<w:name w:val="Footnote Text Char"/>'
            '</w:style>'
            '</w:styles>')


VARIANTS = [
    {"label": "V1_no_styles",                  "with_styles": False, "use_pStyle": False, "spacing": None},
    {"label": "V2_styles_spacing_240_auto",    "with_styles": True,  "use_pStyle": True,  "spacing": 'w:line="240" w:lineRule="auto"'},
    {"label": "V3_styles_spacing_200_auto",    "with_styles": True,  "use_pStyle": True,  "spacing": 'w:line="200" w:lineRule="auto"'},
    {"label": "V4_styles_spacing_200_exact",   "with_styles": True,  "use_pStyle": True,  "spacing": 'w:line="200" w:lineRule="exact"'},
    {"label": "V5_styles_no_spacing",          "with_styles": True,  "use_pStyle": True,  "spacing": None},
    {"label": "V6_styles_spacing_400_atLeast", "with_styles": True,  "use_pStyle": True,  "spacing": 'w:line="400" w:lineRule="atLeast"'},
    {"label": "V7_styles_spacing_350_atLeast", "with_styles": True,  "use_pStyle": True,  "spacing": 'w:line="350" w:lineRule="atLeast"'},
    # V8: with styles.xml but pStyle NOT applied to footnote pPr — does Word auto-apply?
    {"label": "V8_styles_no_pStyleRef",        "with_styles": True,  "use_pStyle": False, "spacing": 'w:line="240" w:lineRule="auto"'},
]


def build(path, variant):
    parts = {
        "[Content_Types].xml": content_types(variant["with_styles"]),
        "_rels/.rels": PKG_RELS,
        "word/_rels/document.xml.rels": doc_rels(variant["with_styles"]),
        "word/document.xml": doc_xml(),
        "word/footnotes.xml": fn_xml(variant["use_pStyle"]),
    }
    if variant["with_styles"]:
        parts["word/styles.xml"] = styles_xml(variant["spacing"])
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in parts.items():
            z.writestr(name, content)


def measure(word, path):
    last = None
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
            time.sleep(0.3)
            fns = doc.Footnotes
            ys = []
            for i in range(1, fns.Count + 1):
                ys.append(round(fns(i).Range.Information(6), 3))
            doc.Close(False)
            return ys
        except Exception as e:
            last = e
            time.sleep(0.5 + attempt * 0.5)
    raise last


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    idx = 0
    try:
        for variant in VARIANTS:
            idx += 1
            path = TMP_DIR / f"ffo_{idx:03d}_{uuid.uuid4().hex[:8]}.docx"
            rec = {**variant}
            try:
                build(path, variant)
                ys = measure(word, path)
                rec["fn_ys"] = ys
                if len(ys) >= 2:
                    rec["lh_implied"] = round(ys[1] - ys[0], 3)
                print(f"[{idx}] {variant['label']:<32} ys={ys} lh={rec.get('lh_implied','-')}")
            except Exception as e:
                rec["error"] = str(e)
                print(f"[{idx}] {variant['label']}: ERR {e}")
            try: path.unlink()
            except: pass
            results.append(rec)
    finally:
        try: word.Quit()
        except: pass
        for f in TMP_DIR.glob("*.docx"):
            try: f.unlink()
            except: pass

    OUT.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
