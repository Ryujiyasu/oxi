"""§4.6.3 — Determine whether w:hint="eastAsia" substitutes for missing
w:eastAsia attribute in triggering CJK-adjacent space widening.

Per spec line 569-585: CJK-adjacent space widening happens "only when the
run's <w:rFonts> has an explicit w:eastAsia attribute". theme-fallback
eastAsiaTheme does NOT count. w:hint="eastAsia" behavior is "untested".

This script generates 5 docx variants by writing OOXML directly, then
opens each via Word COM and measures the advance of the SPACE character
in `Foo は` (Latin foo + space + CJK は).

Variants:
  V1 — ascii + hAnsi only (no eastAsia, no hint, no theme)
  V2 — ascii + hAnsi + eastAsia (explicit)
  V3 — ascii + hAnsi + eastAsiaTheme (theme)
  V4 — ascii + hAnsi + hint="eastAsia" (no explicit eastAsia)
  V5 — ascii + hAnsi + hint="default" (negative ctrl)
  V6 — eastAsia only on the SPACE run (vs neighboring Latin runs ascii-only)

Output: pipeline_data/ra_manual_measurements.json key
"easia_hint_attribute_2026-05-02"
"""
import win32com.client
import os
import time
import json
import zipfile
import shutil
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/easia_hint_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")

DOCUMENT_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:rPr/></w:pPr>
      <w:r>
        <w:rPr>__RPR__</w:rPr>
        <w:t xml:space="preserve">__TEXT__</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
    </w:sectPr>
  </w:body>
</w:document>
"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/theme/theme1.xml"
   ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>
"""

RELS_ROOT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

WORD_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>
"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
    <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
  </w:compat>
</w:settings>
"""

# Minimal theme1.xml using built-in MS minorHAnsi=Calibri, minorEastAsia=ＭＳ 明朝
THEME_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/><a:cs typeface=""/>
        <a:font script="Jpan" typeface="ＭＳ 明朝"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/><a:cs typeface=""/>
        <a:font script="Jpan" typeface="ＭＳ 明朝"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln><a:noFill/></a:ln><a:ln><a:noFill/></a:ln><a:ln><a:noFill/></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
"""

# Variants
VARIANTS = [
    ("V1_no_easia",
     '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
     '<w:sz w:val="24"/>'),
    ("V2_explicit_easia",
     '<w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman"/>'
     '<w:sz w:val="24"/>'),
    ("V3_easia_theme",
     '<w:rFonts w:ascii="Times New Roman" w:eastAsiaTheme="minorEastAsia" w:hAnsi="Times New Roman"/>'
     '<w:sz w:val="24"/>'),
    ("V4_hint_eastAsia",
     '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="eastAsia"/>'
     '<w:sz w:val="24"/>'),
    ("V5_hint_default",
     '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:hint="default"/>'
     '<w:sz w:val="24"/>'),
    ("V6_explicit_easia_with_hint",
     '<w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman" w:hint="eastAsia"/>'
     '<w:sz w:val="24"/>'),
    ("V7_explicit_multi_easia",
     '<w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ 明朝, Yu Mincho" w:hAnsi="Times New Roman"/>'
     '<w:sz w:val="24"/>'),
]

PROBE_TEXT = "Foo は M"


def write_docx(out_path: str, rpr_inner: str, text: str):
    tmp = tempfile.mkdtemp(prefix="easia_doc_")
    try:
        os.makedirs(os.path.join(tmp, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "word", "_rels"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "word", "theme"), exist_ok=True)
        with open(os.path.join(tmp, "[Content_Types].xml"), "w",
                  encoding="utf-8") as f:
            f.write(CONTENT_TYPES)
        with open(os.path.join(tmp, "_rels", ".rels"), "w",
                  encoding="utf-8") as f:
            f.write(RELS_ROOT)
        with open(os.path.join(tmp, "word", "_rels", "document.xml.rels"),
                  "w", encoding="utf-8") as f:
            f.write(WORD_RELS)
        with open(os.path.join(tmp, "word", "styles.xml"), "w",
                  encoding="utf-8") as f:
            f.write(STYLES_XML)
        with open(os.path.join(tmp, "word", "settings.xml"), "w",
                  encoding="utf-8") as f:
            f.write(SETTINGS_XML)
        with open(os.path.join(tmp, "word", "theme", "theme1.xml"), "w",
                  encoding="utf-8") as f:
            f.write(THEME_XML)
        doc_xml = (DOCUMENT_XML_TEMPLATE
                   .replace("__RPR__", rpr_inner)
                   .replace("__TEXT__", text))
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(doc_xml)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def measure_advances(word, docx_path):
    d = word.Documents.Open(docx_path, ReadOnly=True)
    time.sleep(0.20)
    chars = d.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            xs.append((ch,
                       float(c.Information(5)),
                       c.Font.Name,
                       c.Font.Size))
        except Exception:
            continue
    d.Close(SaveChanges=False)
    advs = []
    for i in range(len(xs) - 1):
        advs.append({
            "ch": xs[i][0],
            "adv": round(xs[i + 1][1] - xs[i][1], 4),
            "font": xs[i][2],
            "size": xs[i][3],
        })
    advs.append({"ch": xs[-1][0], "adv": None,
                 "font": xs[-1][2], "size": xs[-1][3]})
    return advs


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for label, rpr_inner in VARIANTS:
            doc_path = os.path.join(OUT_DIR, f"{label}.docx")
            write_docx(doc_path, rpr_inner, PROBE_TEXT)
            try:
                advs = measure_advances(word, doc_path)
            except Exception as e:
                advs = {"error": str(e)}
            results[label] = {"text": PROBE_TEXT, "advances": advs}
            print(f"[{label}] {advs}", flush=True)
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    if os.path.exists(RESULT_PATH):
        try:
            with open(RESULT_PATH, encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = {}
    else:
        existing = {}
    existing["easia_hint_attribute_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
