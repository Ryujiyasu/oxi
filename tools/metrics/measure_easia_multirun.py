"""§4.6.3 follow-up — test multi-run hypothesis.

Hypothesis: CJK-adjacent space widening triggers when Latin and CJK
chars are in separate runs (with run-boundary between them), regardless
of rFonts attributes within each run.

Variants:
  M1: single run with all chars (V1 baseline)
  M2: split into Latin run + CJK run + Latin run (3 runs)
  M3: same 3 runs but with explicit eastAsia on CJK run only
  M4: same 3 runs with explicit eastAsia on ALL runs
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

OUT_DIR = os.path.abspath("pipeline_data/easia_multirun_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")

# Common boilerplate (subset of measure_easia_hint_attribute.py)
DOC_HEADER = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:rPr/></w:pPr>'''
DOC_FOOTER = '''    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
'''

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
</Relationships>
"""
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
</w:styles>
"""
SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>
"""


def write_docx(out_path, runs_xml):
    tmp = tempfile.mkdtemp(prefix="multirun_doc_")
    try:
        os.makedirs(os.path.join(tmp, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "word", "_rels"), exist_ok=True)
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
        doc_xml = DOC_HEADER + runs_xml + DOC_FOOTER
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


def runs_for_variant(variant: str) -> str:
    rpr_no = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
    rpr_ea = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'

    if variant == "M1_single_run":
        return f'<w:r>{rpr_no}<w:t xml:space="preserve">Foo は M</w:t></w:r>'

    if variant == "M2_3runs_no_easia":
        return (f'<w:r>{rpr_no}<w:t xml:space="preserve">Foo </w:t></w:r>'
                f'<w:r>{rpr_no}<w:t>は</w:t></w:r>'
                f'<w:r>{rpr_no}<w:t xml:space="preserve"> M</w:t></w:r>')

    if variant == "M3_3runs_easia_only_cjk":
        return (f'<w:r>{rpr_no}<w:t xml:space="preserve">Foo </w:t></w:r>'
                f'<w:r>{rpr_ea}<w:t>は</w:t></w:r>'
                f'<w:r>{rpr_no}<w:t xml:space="preserve"> M</w:t></w:r>')

    if variant == "M4_3runs_easia_all":
        return (f'<w:r>{rpr_ea}<w:t xml:space="preserve">Foo </w:t></w:r>'
                f'<w:r>{rpr_ea}<w:t>は</w:t></w:r>'
                f'<w:r>{rpr_ea}<w:t xml:space="preserve"> M</w:t></w:r>')

    if variant == "M5_5runs_easia_only_cjk":
        # Foo / space / は / space / M (5 runs)
        return (f'<w:r>{rpr_no}<w:t>Foo</w:t></w:r>'
                f'<w:r>{rpr_no}<w:t xml:space="preserve"> </w:t></w:r>'
                f'<w:r>{rpr_ea}<w:t>は</w:t></w:r>'
                f'<w:r>{rpr_no}<w:t xml:space="preserve"> </w:t></w:r>'
                f'<w:r>{rpr_no}<w:t>M</w:t></w:r>')

    if variant == "M6_5runs_space_in_cjk_run":
        # Foo / [space は space] / M (3 runs but spaces in CJK run)
        return (f'<w:r>{rpr_no}<w:t>Foo</w:t></w:r>'
                f'<w:r>{rpr_ea}<w:t xml:space="preserve"> は </w:t></w:r>'
                f'<w:r>{rpr_no}<w:t>M</w:t></w:r>')

    raise ValueError(variant)


VARIANTS = [
    "M1_single_run",
    "M2_3runs_no_easia",
    "M3_3runs_easia_only_cjk",
    "M4_3runs_easia_all",
    "M5_5runs_easia_only_cjk",
    "M6_5runs_space_in_cjk_run",
]


def measure(word, path):
    d = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.2)
    chars = d.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            t = c.Text
            if t in ("\r", "\x07"):
                continue
            xs.append((t, float(c.Information(5)),
                       c.Font.Name, c.Font.Size))
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
    return advs


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for variant in VARIANTS:
            path = os.path.join(OUT_DIR, f"{variant}.docx")
            runs_xml = runs_for_variant(variant)
            write_docx(path, runs_xml)
            try:
                advs = measure(word, path)
            except Exception as e:
                advs = {"error": str(e)}
            results[variant] = advs
            print(f"[{variant}] {advs}", flush=True)
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
    existing["easia_multirun_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
