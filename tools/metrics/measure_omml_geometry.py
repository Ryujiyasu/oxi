"""Per-glyph geometry measurement for OMML primitives.

Measures Word's rendering geometry (surrogate-pair-aware):
- 01_frac: a/b — num x,y; den x,y; bar y (implicit, between them)
- 02_sup: x^2 — base x,y; sup x,y (offset above)
- 03_sub: x_1 — base x,y; sub x,y (offset below)

Each fixture uses unique char per slot (A/B, X/2, Y/1) for unambiguous ID.

Output: tools/metrics/output/omml_geometry.json
"""
import json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP = Path("pipeline_data") / "_omml_geom_tmp.docx"
OUT = Path(__file__).with_name("output") / "omml_geometry.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def build_docx(omml_body):
    """omml_body is a sequence of <w:p> elements."""
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}>'
        f'<w:body>{omml_body}'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def read_chars(word):
    """Iterate math paragraph per-char with surrogate-pair aggregation.
    Returns list of {ch, cp, x, y} for each glyph-like position.
    """
    doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
    time.sleep(0.5)
    # Find the math paragraph — first oMathPara (paragraph 2 typically after label)
    rng = None
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        # Does the para contain any OMML? Check presence of math in the Range.Text (rendered chars)
        text = p.Range.Text or ""
        has_math_italic = any(0x1D400 <= ord(c) <= 0x1D7FF for c in text)
        has_math_greek = any(0x1D6A8 <= ord(c) <= 0x1D7CB for c in text)
        if has_math_italic or has_math_greek or '\u210E' in text:
            rng = p.Range
            break

    if rng is None:
        # Fall back to paragraph 2
        rng = doc.Paragraphs(2).Range

    sel = word.Selection
    rendered = []
    ci = rng.Start
    while ci < rng.End:
        sel.SetRange(ci, ci + 1)
        ch = sel.Text
        if not ch:
            ci += 1
            continue
        if ch in ('\r', '\x07', '\n'):
            ci += 1
            continue
        if 0xD800 <= ord(ch[0]) <= 0xDBFF:
            sel.SetRange(ci + 1, ci + 2)
            low = sel.Text
            if low and 0xDC00 <= ord(low[0]) <= 0xDFFF:
                cp = 0x10000 + (ord(ch[0]) - 0xD800) * 0x400 + (ord(low[0]) - 0xDC00)
                # Get x,y at this position
                sel.SetRange(ci, ci + 2)
                x = sel.Information(5)
                y = sel.Information(6)
                rendered.append({"ch": ch + low, "cp": cp, "x": round(x, 2), "y": round(y, 2)})
                ci += 2
                continue
        x = sel.Information(5)
        y = sel.Information(6)
        rendered.append({"ch": ch, "cp": ord(ch[0]), "x": round(x, 2), "y": round(y, 2)})
        ci += 1
    doc.Close(False)
    return rendered


FIXTURES = {
    "fraction_A_B": {
        "desc": "fraction A/B",
        "math": '<m:f><m:num><m:r><m:t>A</m:t></m:r></m:num>'
                '<m:den><m:r><m:t>B</m:t></m:r></m:den></m:f>',
    },
    "superscript_X_2": {
        "desc": "X^2 (superscript)",
        "math": '<m:sSup><m:e><m:r><m:t>X</m:t></m:r></m:e>'
                '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
    },
    "subscript_Y_1": {
        "desc": "Y_1 (subscript)",
        "math": '<m:sSub><m:e><m:r><m:t>Y</m:t></m:r></m:e>'
                '<m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub>',
    },
}


def main():
    word = None
    for attempt in range(5):
        try:
            word = win32com.client.Dispatch("Word.Application")
            time.sleep(2.0)
            word.Visible = False
            word.DisplayAlerts = False
            break
        except Exception as e:
            print(f"  Word COM attempt {attempt+1} err: {e}")
            time.sleep(10 * (attempt + 1))
    if word is None:
        return

    results = {}
    try:
        for name, spec in FIXTURES.items():
            print(f"\n=== {name} ===")
            body = (
                '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BEFORE</w:t></w:r></w:p>'
                '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
                '<m:oMathPara><m:oMath>'
                f'{spec["math"]}'
                '</m:oMath></m:oMathPara></w:p>'
                '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>AFTER</w:t></w:r></w:p>'
            )
            build_docx(body)
            try:
                glyphs = read_chars(word)
            except Exception as e:
                print(f"  err: {e}")
                results[name] = {"error": str(e)}
                continue
            print(f"  {len(glyphs)} glyphs measured:")
            for g in glyphs:
                cp = g["cp"]
                print(f"    cp=U+{cp:04X} x={g['x']:>7.2f} y={g['y']:>7.2f}")
            # Analysis specific to fixture
            if len(glyphs) >= 2:
                a = glyphs[0]; b = glyphs[1]
                dy = round(b["y"] - a["y"], 2)
                dx = round(b["x"] - a["x"], 2)
                print(f"  glyph1 → glyph2:  dy={dy:+} dx={dx:+}")
                results[name] = {
                    "desc": spec["desc"],
                    "glyphs": glyphs,
                    "dy_g1_g2": dy,
                    "dx_g1_g2": dx,
                }
            else:
                results[name] = {"desc": spec["desc"], "glyphs": glyphs}
    finally:
        try: word.Quit()
        except: pass
        try: os.remove(TMP)
        except: pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=True, indent=2)
    print(f"\nSaved → {OUT}")


if __name__ == "__main__":
    main()
