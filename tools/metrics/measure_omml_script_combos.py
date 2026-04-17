"""Measure script combinations: sSub, sSup, sSubSup, sPre.

- sSup: x^2
- sSub: x_1
- sSubSup: x^2_1 (both at once) — tests SubSuperscriptGapMin
- sPre: ^14_6 C (pre-scripts on isotope-style notation)
- Nested: (x^2)^3 — tests script-size cascade (73% then 60%)

Measures per-char x positions (y limited per prior finding).
Output: tools/metrics/output/omml_script_combos.json
"""
import json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP = Path("pipeline_data") / "_omml_script_tmp.docx"
OUT = Path(__file__).with_name("output") / "omml_script_combos.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def build_docx(math_frag, display_mode=False):
    wrapper = f'<m:oMathPara><m:oMath>{math_frag}</m:oMath></m:oMathPara>' if display_mode else f'<m:oMath>{math_frag}</m:oMath>'
    body = (
        '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BEFORE</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>{wrapper}</w:p>'
        '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>AFTER</w:t></w:r></w:p>'
    )
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}>'
        f'<w:body>{body}'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>'
    )
    for _ in range(3):
        try:
            if TMP.exists():
                os.remove(TMP)
            break
        except PermissionError:
            time.sleep(0.5)
    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def read_glyphs(word):
    doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
    time.sleep(0.3)
    rng = doc.Paragraphs(2).Range  # math paragraph
    sel = word.Selection
    glyphs = []
    ci = rng.Start
    while ci < rng.End:
        sel.SetRange(ci, ci + 1)
        ch = sel.Text
        if not ch or ch in ('\r', '\x07', '\n'):
            ci += 1
            continue
        if 0xD800 <= ord(ch[0]) <= 0xDBFF:
            sel.SetRange(ci + 1, ci + 2)
            low = sel.Text
            if low and 0xDC00 <= ord(low[0]) <= 0xDFFF:
                cp = 0x10000 + (ord(ch[0]) - 0xD800) * 0x400 + (ord(low[0]) - 0xDC00)
                sel.SetRange(ci, ci + 2)
                x = sel.Information(5); y = sel.Information(6)
                glyphs.append({"cp": cp, "ch": ch + low, "x": round(x, 2), "y": round(y, 2)})
                ci += 2
                continue
        x = sel.Information(5); y = sel.Information(6)
        glyphs.append({"cp": ord(ch[0]), "ch": ch, "x": round(x, 2), "y": round(y, 2)})
        ci += 1
    # Also capture para heights
    y1 = doc.Paragraphs(1).Range.Information(6)
    y2 = doc.Paragraphs(2).Range.Information(6)
    try:
        y3 = doc.Paragraphs(3).Range.Information(6)
    except Exception:
        y3 = None
    doc.Close(False)
    return {"glyphs": glyphs, "y1": round(y1, 2), "y2": round(y2, 2), "y3": round(y3, 2) if y3 else None}


TESTS = [
    # label, math, display_mode
    ("sSup_x2",      '<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>', False),
    ("sSub_x1",      '<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub>', False),
    ("sSubSup_x12",  '<m:sSubSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>1</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>', False),
    ("sPre_14_6_C",  '<m:sPre><m:sub><m:r><m:t>6</m:t></m:r></m:sub><m:sup><m:r><m:t>14</m:t></m:r></m:sup><m:e><m:r><m:t>C</m:t></m:r></m:e></m:sPre>', False),
    ("sSup_nested",  '<m:sSup><m:e><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>a</m:t></m:r></m:sup></m:sSup></m:e><m:sup><m:r><m:t>b</m:t></m:r></m:sup></m:sSup>', False),
]


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
            print(f"  attempt {attempt+1}: {e}")
            time.sleep(10 * (attempt + 1))
    if word is None:
        return

    results = []
    try:
        for label, math, display in TESTS:
            print(f"\n=== {label} ===")
            build_docx(math, display_mode=display)
            try:
                data = read_glyphs(word)
            except Exception as e:
                print(f"  ERR: {e}")
                results.append({"label": label, "error": str(e)})
                continue
            print(f"  y_before={data['y1']}  y_math={data['y2']}  y_after={data['y3']}")
            line_h = data['y3'] - data['y2'] if data['y3'] else None
            print(f"  line_h={line_h}")
            print(f"  {len(data['glyphs'])} glyphs:")
            for g in data['glyphs']:
                print(f"    cp=U+{g['cp']:04X}  x={g['x']:>7.2f}  y={g['y']:>7.2f}  ch={g['ch']!r}")
            # Horizontal differences for script analysis
            if len(data['glyphs']) >= 2:
                dxs = [round(data['glyphs'][i+1]['x'] - data['glyphs'][i]['x'], 2)
                       for i in range(len(data['glyphs'])-1)]
                print(f"  dx chain: {dxs}")
            results.append({"label": label, "data": data, "line_h": line_h})
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
