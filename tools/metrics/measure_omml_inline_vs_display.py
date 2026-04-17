"""Measure inline vs display fraction heights to confirm MATH table
constant selection (FractionNumeratorShiftUp vs FractionNumeratorDisplayStyleShiftUp).

The same fraction A/B renders differently in:
  - Inline context (<m:oMath> directly in running text) — uses inline constants
  - Display context (<m:oMathPara>) — uses display-style constants

Expected difference (from Cambria Math MATH table, 10.5pt scaled):
  inline FractionNumeratorShiftUp:   1200 / 2048 * 10.5 = 6.15pt
  display FractionNumeratorShiftUp:  1550 / 2048 * 10.5 = 7.95pt
  Difference ≈ 1.8pt taller for display.

Also expected:
  inline DenominatorShiftDown: 1030 / 2048 * 10.5 = 5.28pt
  display DenominatorShiftDown: 1370 / 2048 * 10.5 = 7.03pt
  Combined display is ~3.5pt taller than inline for same fraction.

Output: tools/metrics/output/omml_inline_vs_display.json
"""
import json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP = Path("pipeline_data") / "_omml_inline_tmp.docx"
OUT = Path(__file__).with_name("output") / "omml_inline_vs_display.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def build_docx(body_paras):
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}>'
        f'<w:body>{body_paras}'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


FRACTION_OMML = '<m:f><m:num><m:r><m:t>A</m:t></m:r></m:num><m:den><m:r><m:t>B</m:t></m:r></m:den></m:f>'

# Test 1: Inline math inside body text
INLINE_BODY = (
    '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
    '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BEFORE </w:t></w:r>'
    f'<m:oMath>{FRACTION_OMML}</m:oMath>'
    '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t> AFTER</w:t></w:r>'
    '</w:p>'
    '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>NEXT</w:t></w:r></w:p>'
)

# Test 2: Display math (oMathPara)
DISPLAY_BODY = (
    '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BEFORE</w:t></w:r></w:p>'
    '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
    f'<m:oMathPara><m:oMath>{FRACTION_OMML}</m:oMath></m:oMathPara>'
    '</w:p>'
    '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>NEXT</w:t></w:r></w:p>'
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
            text = p.Range.Text[:30]
            ys.append({"idx": i, "y": round(y, 2), "text": text})
        doc.Close(False)

        # Find the "math paragraph" (contains math-italic char) and "next" paragraph
        # For inline: math is in para 1 (with BEFORE...AFTER), next is para 2
        # For display: math is in para 2 (alone), next is para 3

        result = {"label": label, "para_ys": ys}
        if label == "inline":
            # Math in para 1; line height extended vs pure text?
            # Compare: render same doc with pure text in para 1 to get baseline
            pass
        elif label == "display":
            # Math height = y[3] - y[2]
            if len(ys) >= 3:
                result["math_para_height"] = round(ys[2]["y"] - ys[1]["y"], 2)
        return result
    except Exception as e:
        return {"label": label, "error": str(e)}


def measure_pure_text_baseline(word):
    """Baseline: same layout without math to get pure text line-height."""
    body = (
        '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BEFORE AFTER</w:t></w:r></w:p>'
        '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>NEXT</w:t></w:r></w:p>'
    )
    build_docx(body)
    try:
        doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
        time.sleep(0.5)
        y1 = doc.Paragraphs(1).Range.Information(6)
        y2 = doc.Paragraphs(2).Range.Information(6)
        doc.Close(False)
        return {"baseline_line_h": round(y2 - y1, 2), "y1": round(y1, 2), "y2": round(y2, 2)}
    except Exception as e:
        return {"error": str(e)}


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

    try:
        print("=== Pure text baseline (no math) ===")
        baseline = measure_pure_text_baseline(word)
        print(f"  {baseline}")

        print("\n=== Inline math (BEFORE A/B AFTER in single paragraph) ===")
        inline = measure(word, INLINE_BODY, "inline")
        for p in inline.get("para_ys", []):
            print(f"  para{p['idx']}: y={p['y']:>7.2f} text={p['text']!r}")

        print("\n=== Display math (oMathPara alone) ===")
        display = measure(word, DISPLAY_BODY, "display")
        for p in display.get("para_ys", []):
            print(f"  para{p['idx']}: y={p['y']:>7.2f} text={p['text']!r}")
        if "math_para_height" in display:
            print(f"  display math_para_height: {display['math_para_height']}")

        # Analysis
        print("\n=== Analysis ===")
        # Inline: if math is in para 1, next para y2 = y1 + line_h_with_math.
        # line_h should be taller than baseline if fraction grew the line.
        if len(inline.get("para_ys", [])) >= 2:
            inline_line_h = inline["para_ys"][1]["y"] - inline["para_ys"][0]["y"]
            print(f"  Inline line height (math in para 1): {round(inline_line_h, 2)}")
            if "baseline_line_h" in baseline:
                print(f"  Baseline line height (no math):      {baseline['baseline_line_h']}")
                print(f"  Diff (inline math adds):             {round(inline_line_h - baseline['baseline_line_h'], 2)}")
        if "math_para_height" in display:
            print(f"  Display math paragraph height:       {display['math_para_height']}")

        out = {
            "baseline": baseline,
            "inline": inline,
            "display": display,
        }
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=True, indent=2)
        print(f"\nSaved → {OUT}")
    finally:
        try: word.Quit()
        except: pass
        try: os.remove(TMP)
        except: pass


if __name__ == "__main__":
    main()
