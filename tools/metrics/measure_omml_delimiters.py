"""Measure delimiter growth behavior for <m:d> (brackets around content).

Tests:
  - (x)         simple, small parens
  - (a/b)       fraction inside — parens grow
  - (a^2+b^2)   superscript + binary op — parens grow moderately
  - [a/b]       square brackets
  - |x/y|       absolute value bars
  - {x}         curly braces
  - ((x/y))     nested parens

Output: tools/metrics/output/omml_delimiters.json
"""
import json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP = Path("pipeline_data") / "_omml_delim_tmp.docx"
OUT = Path(__file__).with_name("output") / "omml_delimiters.json"
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


def delim(beg, end, content):
    """Build m:d with given delimiters around content (inner OMML)."""
    return (
        '<m:d><m:dPr>'
        f'<m:begChr m:val="{beg}"/>'
        f'<m:endChr m:val="{end}"/>'
        '</m:dPr>'
        f'<m:e>{content}</m:e>'
        '</m:d>'
    )


def read(word):
    doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
    time.sleep(0.3)
    ys = []
    for i in range(1, doc.Paragraphs.Count + 1):
        try:
            y = doc.Paragraphs(i).Range.Information(6)
            ys.append(round(y, 2))
        except Exception:
            ys.append(None)
    doc.Close(False)
    return ys


frac = '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>'
sup_ab = '<m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>+</m:t></m:r><m:sSup><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>'
sqrt_x = '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>'

TESTS = [
    ("paren_x",         delim("(", ")", '<m:r><m:t>x</m:t></m:r>')),
    ("paren_frac",      delim("(", ")", frac)),
    ("paren_sup_ab",    delim("(", ")", sup_ab)),
    ("paren_sqrt",      delim("(", ")", sqrt_x)),
    ("bracket_frac",    delim("[", "]", frac)),
    ("abs_frac",        delim("|", "|", frac)),
    ("brace_frac",      delim("{", "}", frac)),
    ("nested_paren",    delim("(", ")", delim("(", ")", frac))),
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

    # Baseline: plain fraction without delimiter
    print("=== baseline: plain fraction a/b inline ===")
    build_docx(frac, display_mode=False)
    try:
        ys = read(word)
        if len(ys) >= 3:
            baseline_h = round(ys[2] - ys[1], 2)
            print(f"  baseline math_h = {baseline_h}")
        else:
            baseline_h = None
    except Exception as e:
        print(f"  ERR: {e}")
        baseline_h = None

    results = {"baseline_frac_inline": baseline_h}
    try:
        for label, math in TESTS:
            time.sleep(0.3)
            build_docx(math, display_mode=False)
            try:
                ys = read(word)
            except Exception as e:
                print(f"  {label}: ERR {e}")
                results[label] = {"error": str(e)}
                continue
            math_h = round(ys[2] - ys[1], 2) if len(ys) >= 3 and ys[2] and ys[1] else None
            delta = round(math_h - baseline_h, 2) if math_h and baseline_h else None
            print(f"  {label:<20} math_h={math_h}  Δ_vs_baseline={delta}")
            results[label] = {"math_h": math_h, "delta_vs_baseline": delta}
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
