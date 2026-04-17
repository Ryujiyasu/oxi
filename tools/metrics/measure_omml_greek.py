"""Measure Word's substitution of Greek letters in OMML runs.

Tests α-ω (U+03B1-03C9) and Α-Ω (U+0391-03A9). Math Italic Greek range
is U+1D6E2-U+1D71B. Also tests variant forms: ϑ ϕ ϖ ϱ ϵ.

Key question: does Word auto-italicize Greek the same way as Latin,
or keep them as entered? By mathematical convention, variables are
italic but operators (∑, Π, ∏) are upright.

Output: tools/metrics/output/omml_greek_table.json
"""
import json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP = Path("pipeline_data") / "_omml_greek_tmp.docx"
TMP.parent.mkdir(parents=True, exist_ok=True)
OUT = Path(__file__).with_name("output") / "omml_greek_table.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

# Greek letter sets
GREEK_LOWER = ['α','β','γ','δ','ε','ζ','η','θ','ι','κ','λ','μ','ν','ξ','ο','π','ρ','σ','τ','υ','φ','χ','ψ','ω']
GREEK_UPPER = ['Α','Β','Γ','Δ','Ε','Ζ','Η','Θ','Ι','Κ','Λ','Μ','Ν','Ξ','Ο','Π','Ρ','Σ','Τ','Υ','Φ','Χ','Ψ','Ω']
GREEK_VARIANT = ['ϑ','ϕ','ϖ','ϱ','ϵ']  # theta-sym, phi-sym, pi-sym, rho-sym, lunate-epsilon
GREEK_FINAL = ['ς']


def build_docx(letters):
    math_content = ''.join(f'<m:r><m:t>{c}</m:t></m:r>' for c in letters)
    body = (
        '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
        '<m:oMathPara><m:oMath>'
        f'{math_content}'
        '</m:oMath></m:oMathPara></w:p>'
    )
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


def measure(word, letters):
    build_docx(letters)
    try:
        doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
        time.sleep(0.5)
        rng = doc.Paragraphs(1).Range
        sel = word.Selection
        rendered = []
        ci = rng.Start
        while ci < rng.End:
            sel.SetRange(ci, ci + 1)
            ch = sel.Text
            if not ch or ch in ('\r', '\x07', '\n'):
                ci += 1
                continue
            # Handle surrogate pairs
            if 0xD800 <= ord(ch[0]) <= 0xDBFF:
                sel.SetRange(ci + 1, ci + 2)
                low = sel.Text
                if low and 0xDC00 <= ord(low[0]) <= 0xDFFF:
                    cp = 0x10000 + (ord(ch[0]) - 0xD800) * 0x400 + (ord(low[0]) - 0xDC00)
                    rendered.append({"ch": ch + low, "cp": cp})
                    ci += 2
                    continue
            rendered.append({"ch": ch, "cp": ord(ch[0])})
            ci += 1
        doc.Close(False)
        return rendered
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
            print(f"  Word COM attempt {attempt+1} err: {e}")
            time.sleep(10 * (attempt + 1))
    if word is None:
        print("Word COM unavailable. Bailing.")
        return

    all_groups = [
        ("lowercase", GREEK_LOWER),
        ("uppercase", GREEK_UPPER),
        ("variants", GREEK_VARIANT),
        ("final_sigma", GREEK_FINAL),
    ]

    results = {}
    try:
        for name, letters in all_groups:
            print(f"\n=== {name} ({len(letters)} chars) ===")
            r = measure(word, letters)
            if isinstance(r, dict) and "error" in r:
                print(f"  ERR: {r['error']}")
                results[name] = {"error": r['error']}
                continue
            if len(r) != len(letters):
                print(f"  MISMATCH: {len(r)} rendered vs {len(letters)} input")
                results[name] = {"mismatch": True, "rendered": r, "expected": len(letters)}
                continue
            table = []
            for i, c in enumerate(letters):
                entry = {
                    "input_ch": c,
                    "input_cp": ord(c),
                    "rendered_ch": r[i]["ch"],
                    "rendered_cp": r[i]["cp"],
                    "substituted": r[i]["cp"] != ord(c),
                }
                table.append(entry)
                mark = "→" if entry["substituted"] else " "
                print(f"  '{c}' (U+{entry['input_cp']:04X}) {mark} U+{entry['rendered_cp']:04X}")
            results[name] = table
    finally:
        try: word.Quit()
        except: pass
        try: os.remove(TMP)
        except: pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=True, indent=2)

    # Summary
    print(f"\n=== Summary ===")
    for name, entries in results.items():
        if isinstance(entries, list):
            subst = sum(1 for e in entries if e["substituted"])
            print(f"  {name}: {subst}/{len(entries)} substituted")
    print(f"Saved → {OUT}")


if __name__ == "__main__":
    main()
