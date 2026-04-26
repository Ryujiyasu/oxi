"""Measure grow operator behavior (∑ ∫ ∏) in inline vs display contexts.

Tests:
  - Inline ∑: <m:nary><m:chr m:val="∑"/> ...</m:nary> in body
  - Display ∑: same in <m:oMathPara>
  - Limit location: limLoc=subSup (inline style, default) vs undOvr (display)
  - Content height effect: operand = 'x' vs 'x/y' fraction

Compares paragraph heights to infer when Word selects the "grow" glyph
variant from MATH table's 56 vertical variants.

Output: tools/metrics/output/omml_nary_grow.json
"""
import json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP = Path("pipeline_data") / "_omml_nary_tmp.docx"
OUT = Path(__file__).with_name("output") / "omml_nary_grow.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def build_docx(body):
    # Release any existing file
    for _ in range(3):
        try:
            if TMP.exists():
                os.remove(TMP)
            break
        except PermissionError:
            time.sleep(0.5)
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}>'
        f'<w:body>{body}'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def nary_xml(chr_val, lim_loc, sub_text, sup_text, operand_xml):
    return (
        '<m:nary>'
        '<m:naryPr>'
        f'<m:chr m:val="{chr_val}"/>'
        f'<m:limLoc m:val="{lim_loc}"/>'
        '<m:grow m:val="1"/>'
        '</m:naryPr>'
        f'<m:sub><m:r><m:t>{sub_text}</m:t></m:r></m:sub>'
        f'<m:sup><m:r><m:t>{sup_text}</m:t></m:r></m:sup>'
        f'<m:e>{operand_xml}</m:e>'
        '</m:nary>'
    )


def measure(word, body, label):
    build_docx(body)
    try:
        doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
        time.sleep(0.5)
        n_paras = doc.Paragraphs.Count
        ys = []
        for i in range(1, n_paras + 1):
            p = doc.Paragraphs(i)
            y = p.Range.Information(6)
            text = p.Range.Text[:40]
            ys.append({"idx": i, "y": round(y, 2), "text": text})
        doc.Close(False)
        # Math in para 2 (between BEFORE / AFTER)
        math_h = None
        if len(ys) >= 3:
            math_h = round(ys[2]["y"] - ys[1]["y"], 2)
        return {"label": label, "para_ys": ys, "math_h": math_h}
    except Exception as e:
        return {"label": label, "error": str(e)}


def make_test(nary_block, inline_or_display):
    """Build doc body. inline_or_display: 'inline' wraps in <m:oMath>, 'display' wraps in <m:oMathPara>."""
    if inline_or_display == "inline":
        math_wrapper = f'<m:oMath>{nary_block}</m:oMath>'
    else:
        math_wrapper = f'<m:oMathPara>{nary_block}</m:oMathPara>'
    return (
        '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BEFORE</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>{math_wrapper}</w:p>'
        '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>AFTER</w:t></w:r></w:p>'
    )


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

    # Operands: simple x vs a fraction (tall)
    simple_operand = '<m:r><m:t>x</m:t></m:r>'
    tall_operand = '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>'

    # 8 test cases: {∑,∫} × {inline,display} × {simple,tall operand}
    tests = []
    for chr_name, chr_val in [("sum", "∑"), ("int", "∫")]:
        for mode in ["inline", "display"]:
            # In display mode, default limLoc is undOvr (limits above/below)
            # In inline mode, subSup is typical (limits as sub/sup scripts)
            lim_loc = "undOvr" if mode == "display" else "subSup"
            for op_label, op_xml in [("x", simple_operand), ("a/b", tall_operand)]:
                label = f"{chr_name}_{mode}_{op_label}"
                nary = nary_xml(chr_val, lim_loc, "i=1", "n", op_xml)
                tests.append((label, make_test(nary, mode)))

    results = []
    try:
        for label, body in tests:
            # Tiny delay to let Word release file handle between iterations
            time.sleep(0.3)
            r = measure(word, body, label)
            print(f"  {label:<30} math_h={r.get('math_h')}")
            results.append(r)
    finally:
        try: word.Quit()
        except: pass
        try: os.remove(TMP)
        except: pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=True, indent=2)
    print(f"\nSaved → {OUT}")

    # Analysis
    print("\n=== Summary ===")
    print(f"{'case':<30} {'math_h':>8}")
    for r in results:
        if r.get("math_h") is not None:
            print(f"  {r['label']:<30} {r['math_h']:>8.2f}")


if __name__ == "__main__":
    main()
