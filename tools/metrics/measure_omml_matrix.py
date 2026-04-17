"""Measure OMML matrix cell alignment and row spacing.

Tests:
  - 2x2 matrix with mcJc=center (default)
  - 2x2 matrix with mcJc=left
  - 2x2 matrix with mcJc=right
  - 3x3 matrix (varying content widths: A/BB/CCC)
  - 2x2 with fraction content (tall row)

Measures:
  - Per-cell x position (to infer column alignment)
  - Per-row y position (to infer row gap / BaselineDropMin)
  - Total matrix height

Output: tools/metrics/output/omml_matrix.json
"""
import json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP = Path("pipeline_data") / "_omml_mat_tmp.docx"
OUT = Path(__file__).with_name("output") / "omml_matrix.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def build_docx(math):
    body = (
        '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BEFORE</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr><m:oMathPara><m:oMath>{math}</m:oMath></m:oMathPara></w:p>'
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


def matrix_xml(n_cols, mcJc, rows):
    """Build m:m from rows: list of list of cell-content-OMML strings."""
    mc_pr = f'<m:mc><m:mcPr><m:count m:val="{n_cols}"/><m:mcJc m:val="{mcJc}"/></m:mcPr></m:mc>'
    mr_list = []
    for row_cells in rows:
        cells_xml = ''.join(f'<m:e>{c}</m:e>' for c in row_cells)
        mr_list.append(f'<m:mr>{cells_xml}</m:mr>')
    rows_xml = ''.join(mr_list)
    return f'<m:m><m:mPr><m:mcs>{mc_pr}</m:mcs></m:mPr>{rows_xml}</m:m>'


def measure(word):
    doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
    time.sleep(0.3)
    # Get math paragraph (idx 2) char positions
    rng = doc.Paragraphs(2).Range
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
    y1 = doc.Paragraphs(1).Range.Information(6)
    y2 = doc.Paragraphs(2).Range.Information(6)
    try:
        y3 = doc.Paragraphs(3).Range.Information(6)
    except Exception:
        y3 = None
    doc.Close(False)
    return {"glyphs": glyphs, "y_before": round(y1,2), "y_math": round(y2,2), "y_after": round(y3,2) if y3 else None}


def r(t):
    return f'<m:r><m:t>{t}</m:t></m:r>'


TESTS = [
    ("2x2_center", matrix_xml(2, "center", [[r("a"), r("b")], [r("c"), r("d")]])),
    ("2x2_left",   matrix_xml(2, "left",   [[r("a"), r("b")], [r("c"), r("d")]])),
    ("2x2_right",  matrix_xml(2, "right",  [[r("a"), r("b")], [r("c"), r("d")]])),
    ("3x3_center", matrix_xml(3, "center", [[r("A"), r("BB"), r("CCC")], [r("D"), r("EE"), r("FFF")], [r("G"), r("HH"), r("III")]])),
    ("2x2_fracs",  matrix_xml(2, "center", [
        ['<m:f><m:num>'+r("a")+'</m:num><m:den>'+r("b")+'</m:den></m:f>', r("c")],
        [r("d"), r("e")]
    ])),
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
        for label, math in TESTS:
            time.sleep(0.3)
            build_docx(math)
            try:
                data = measure(word)
            except Exception as e:
                print(f"  {label}: ERR {e}")
                results.append({"label": label, "error": str(e)})
                continue
            math_h = data["y_after"] - data["y_math"] if data["y_after"] else None
            print(f"\n=== {label} ===")
            print(f"  math_h = {round(math_h, 2) if math_h else None}")
            # Filter out \r (ASCII 13) glyphs for cleaner output
            content_glyphs = [g for g in data["glyphs"] if g["cp"] != 0x0D]
            print(f"  {len(content_glyphs)} content glyphs:")
            for g in content_glyphs:
                print(f"    cp=U+{g['cp']:04X}  x={g['x']:>7.2f}  y={g['y']:>7.2f}")
            results.append({"label": label, "math_h": round(math_h, 2) if math_h else None, "glyphs": content_glyphs})
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
