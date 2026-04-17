"""Generate 10 minimal OMML fixture docx files for Word COM measurement.

Each fixture contains a single <m:oMathPara> with a specific math structure,
allowing us to measure Word's rendering of each primitive in isolation.

Output: tools/fixtures/omml_samples/
  01_frac.docx        — fraction  a/b
  02_sup.docx         — superscript x^2
  03_sub.docx         — subscript x_1
  04_sqrt.docx        — square root √x
  05_nary_sum.docx    — summation Σ
  06_matrix_2x2.docx  — 2x2 matrix
  07_acc_hat.docx     — accent (hat) x̂
  08_box.docx         — box around x
  09_bar.docx         — overline bar
  10_limit.docx       — lim x→0
"""
import zipfile, sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = Path(__file__).resolve().parent.parent / "fixtures" / "omml_samples"
OUT_DIR.mkdir(parents=True, exist_ok=True)

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


# OMML fragment templates. Each is wrapped in <m:oMathPara><m:oMath>...</m:oMath></m:oMathPara>.
FIXTURES = {
    "01_frac": {
        "desc": "fraction a/b (default bar type)",
        "math": '<m:f><m:fPr><m:ctrlPr><w:rPr><w:i/></w:rPr></m:ctrlPr></m:fPr>'
                '<m:num><m:r><m:t>a</m:t></m:r></m:num>'
                '<m:den><m:r><m:t>b</m:t></m:r></m:den>'
                '</m:f>',
    },
    "02_sup": {
        "desc": "superscript x^2",
        "math": '<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e>'
                '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
    },
    "03_sub": {
        "desc": "subscript x_1",
        "math": '<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e>'
                '<m:sub><m:r><m:t>1</m:t></m:r></m:sub></m:sSub>',
    },
    "04_sqrt": {
        "desc": "square root of x",
        "math": '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>'
                '<m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>',
    },
    "05_nary_sum": {
        "desc": "summation from i=1 to n of i",
        "math": '<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr>'
                '<m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>'
                '<m:sup><m:r><m:t>n</m:t></m:r></m:sup>'
                '<m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>',
    },
    "06_matrix_2x2": {
        "desc": "2x2 matrix [a b; c d]",
        "math": '<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>'
                '<m:e>'
                '<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="2"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>'
                '<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>'
                '<m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr>'
                '</m:m></m:e></m:d>',
    },
    "07_acc_hat": {
        "desc": "accent: x with circumflex hat",
        "math": '<m:acc><m:accPr><m:chr m:val="̂"/></m:accPr>'
                '<m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>',
    },
    "08_box": {
        "desc": "boxed expression (groupChr box)",
        "math": '<m:box><m:boxPr><m:ctrlPr><w:rPr><w:i/></w:rPr></m:ctrlPr></m:boxPr>'
                '<m:e><m:r><m:t>x+1</m:t></m:r></m:e></m:box>',
    },
    "09_bar": {
        "desc": "overline (bar) on x (no pos = default top)",
        # Simplified: omit m:barPr (defaults to top overline).
        # Previous variant with <m:pos m:val="top"/> crashed Word COM.
        "math": '<m:bar><m:e><m:r><m:t>x</m:t></m:r></m:e></m:bar>',
    },
    "10_limit": {
        "desc": "lower limit: lim_{x→0} f(x)",
        # Fixed: base is m:func with fName='lim' and empty e; lim contains sub.
        # Previous variant put bare text 'lim' in m:e which crashed Word COM.
        "math": '<m:limLow>'
                '<m:e><m:func><m:fName><m:r><m:t>lim</m:t></m:r></m:fName>'
                '<m:e><m:r><m:t>f(x)</m:t></m:r></m:e></m:func></m:e>'
                '<m:lim><m:r><m:t>x→0</m:t></m:r></m:lim>'
                '</m:limLow>',
    },
}


def build_docx(out_path: Path, omml_fragment: str, desc: str):
    # Need font specification for math runs to render in Cambria Math.
    # Word auto-applies Cambria Math for OMML content, but we can also specify.
    # Use oMathPara for display-mode math (standalone paragraph)
    math_para = (
        '<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'
        f'<m:oMath>{omml_fragment}</m:oMath>'
        '</m:oMathPara>'
    )
    # Label paragraph before and after for measurement anchoring
    label_p = f'<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Fixture: {desc}</w:t></w:r></w:p>'
    # OMML paragraph as display math
    math_wrap = (
        '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
        f'{math_para}'
        '</w:p>'
    )
    trailing_p = '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>END</w:t></w:r></w:p>'
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}>'
        f'<w:body>{label_p}{math_wrap}{trailing_p}{sect}</w:body></w:document>'
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def main():
    print(f"Generating {len(FIXTURES)} OMML fixtures → {OUT_DIR}\n")
    for name, spec in FIXTURES.items():
        out = OUT_DIR / f"{name}.docx"
        build_docx(out, spec["math"], spec["desc"])
        print(f"  {name:<20} {spec['desc']}")
    print(f"\n{len(FIXTURES)} fixtures generated.")


if __name__ == "__main__":
    main()
