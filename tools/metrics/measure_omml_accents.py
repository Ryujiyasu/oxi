"""Measure accent positioning on various base chars.

Tests m:acc with common mathematical accents on different base widths:

Accents:
  ̂ (U+0302, circumflex/hat)
  ̃ (U+0303, tilde)
  ̄ (U+0304, macron)
  ̇ (U+0307, dot above)
  ̈ (U+0308, dieresis)
  ⃗ (U+20D7, combining right arrow — vector)

Base chars (narrow to wide):
  i  (narrow, italicized to 𝑖 U+1D456)
  x  (typical, italicized to 𝑥 U+1D465)
  M  (wide, italicized to 𝑀 U+1D440)
  ω  (Greek, italicized to 𝜔 U+1D714)

Expected: Word uses Cambria Math's MATH TopAccentAttachment table
(439 entries) to horizontally center each accent over the base glyph's
optical center.

Output: tools/metrics/output/omml_accents.json
"""
import json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP = Path("pipeline_data") / "_omml_acc_tmp.docx"
OUT = Path(__file__).with_name("output") / "omml_accents.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def build_docx(math):
    body = (
        '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>BEFORE</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr><m:oMath>{math}</m:oMath></w:p>'
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


def acc_xml(accent_char, base_char):
    return (
        '<m:acc>'
        f'<m:accPr><m:chr m:val="{accent_char}"/></m:accPr>'
        f'<m:e><m:r><m:t>{base_char}</m:t></m:r></m:e>'
        '</m:acc>'
    )


def measure(word):
    doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
    time.sleep(0.3)
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
                glyphs.append({"cp": cp, "x": round(x, 2), "y": round(y, 2)})
                ci += 2
                continue
        x = sel.Information(5); y = sel.Information(6)
        glyphs.append({"cp": ord(ch[0]), "x": round(x, 2), "y": round(y, 2)})
        ci += 1
    y1 = doc.Paragraphs(1).Range.Information(6)
    y2 = doc.Paragraphs(2).Range.Information(6)
    y3 = doc.Paragraphs(3).Range.Information(6)
    doc.Close(False)
    return {"glyphs": glyphs, "y_math": round(y2,2), "line_h": round(y3 - y2, 2)}


ACCENTS = [
    ("hat",      '\u0302'),  # ̂
    ("tilde",    '\u0303'),  # ̃
    ("macron",   '\u0304'),  # ̄
    ("dot",      '\u0307'),  # ̇
    ("dieresis", '\u0308'),  # ̈
    ("vector",   '\u20D7'),  # ⃗
]

BASES = [
    ("i_narrow", 'i'),
    ("x",        'x'),
    ("M_wide",   'M'),
    ("omega",    'ω'),
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
        for acc_name, acc_ch in ACCENTS:
            for base_name, base_ch in BASES:
                label = f"{acc_name}_{base_name}"
                math = acc_xml(acc_ch, base_ch)
                time.sleep(0.3)
                build_docx(math)
                try:
                    data = measure(word)
                except Exception as e:
                    print(f"  {label}: ERR {e}")
                    results.append({"label": label, "error": str(e)})
                    continue
                glyphs = [g for g in data["glyphs"] if g["cp"] != 0x0D]
                print(f"{label:<20} line_h={data['line_h']:<5}  glyphs:")
                for g in glyphs:
                    cp = g["cp"]
                    print(f"   cp=U+{cp:04X}  x={g['x']:>7.2f}  y={g['y']:>7.2f}")
                results.append({"label": label, "acc_ch": acc_ch, "base_ch": base_ch,
                                "line_h": data["line_h"], "glyphs": glyphs})
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
