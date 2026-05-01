"""§4.6.2 autoSpaceDE w:val="0" — what does it actually disable?

Per spec §4.6.2: "Active by default; disabled when
<w:autoSpaceDE w:val="0"/> is set."

But the spec only described autoSpaceDE as the kana→Latin alphanumeric
boundary widening. In light of our two-mechanism yakumono finding,
also test:
- Mechanism 1 (Type A/B/C) with autoSpaceDE off
- Mechanism 2 (justify-time) with autoSpaceDE off
- §4.6.2 boundary widening with autoSpaceDE off

Each test pair: same probe with autoSpaceDE on (default) vs explicitly
off via pPr setting.

Build doc via direct OOXML to control pPr.
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

OUT_DIR = os.path.abspath("pipeline_data/autospacede_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")

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

DOC_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:jc w:val="__JC__"/>
__AUTOSPACE__
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman"/>
          <w:sz w:val="24"/>
        </w:rPr>
        <w:t xml:space="preserve">__TEXT__</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="__PGW_TW__" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1700" w:bottom="1440" w:left="1700"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""


def write_docx(path, text, jc, autospacede_off, page_w_pt):
    tmp = tempfile.mkdtemp(prefix="ase_doc_")
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
        autospace_xml = ('<w:autoSpaceDE w:val="0"/>'
                          '<w:autoSpaceDN w:val="0"/>'
                          if autospacede_off else "")
        page_w_tw = int(page_w_pt * 20)
        doc_xml = (DOC_TEMPLATE
                   .replace("__JC__", jc)
                   .replace("__AUTOSPACE__", autospace_xml)
                   .replace("__TEXT__", text)
                   .replace("__PGW_TW__", str(page_w_tw)))
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(doc_xml)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# Test cases
TESTS = [
    # (label, text, jc, page_w_pt — controls overflow for Mech 2)
    # 1. §4.6.2 boundary: 漢M (CJK→Latin) at fixed page
    ("4_6_2_boundary",       "漢M漢M漢",     "left",   500),
    # 2. Mech 1: 漢」（漢 (B→A pair)
    ("mech1_pair",           "漢」（漢",      "left",   500),
    # 3. Mech 1 chain: （（（（
    ("mech1_chain_open",     "（（（（漢",    "left",   500),
    # 4. Mech 2: 27-char text at narrow content
    ("mech2_overflow",
     "漢漢漢「漢漢漢」漢漢漢「漢漢漢」漢漢漢、漢漢漢、漢漢漢",
     "both",
     320),  # narrow — content_w = 320 - 2*85 = ~150pt? let me check
]

# Actually let me make the page widths match what we want for content_w:
# pgMar left=1700tw=85pt, right=1700tw=85pt → content = page_w - 170pt
# For mech2 we want content_w ≈ 320pt, so page_w = 490pt
# For mech1 tests we want content_w large enough no overflow → page_w = 500pt → content=330pt

TESTS = [
    ("4_6_2_boundary",   "漢M漢M漢",                                          "left",  500),
    ("mech1_pair",       "漢」（漢",                                          "left",  500),
    ("mech1_chain_open", "（（（（漢",                                        "left",  500),
    ("mech2_overflow",
     "漢漢漢「漢漢漢」漢漢漢「漢漢漢」漢漢漢、漢漢漢、漢漢漢",
     "both",  490),  # content_w = 320pt
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for label, text, jc, pgW in TESTS:
            results[label] = {}
            for autospace_off in [False, True]:
                key = "off" if autospace_off else "on"
                fname = f"{label}_{key}.docx"
                path = os.path.join(OUT_DIR, fname)
                write_docx(path, text, jc, autospace_off, pgW)
                try:
                    d = word.Documents.Open(path, ReadOnly=True)
                    time.sleep(0.3)
                    chars = d.Range().Characters
                    xs = []
                    for ci in range(1, chars.Count + 1):
                        try:
                            c = chars(ci)
                            t = c.Text
                            if t in ("\r", "\x07"):
                                continue
                            xs.append((t,
                                       float(c.Information(5)),
                                       float(c.Information(6))))
                        except Exception:
                            continue
                    d.Close(SaveChanges=False)
                    if not xs:
                        continue
                    y0 = xs[0][2]
                    line1 = [(c, x) for c, x, y in xs
                             if abs(y - y0) < 0.5]
                    line1_sorted = sorted(line1, key=lambda t: t[1])
                    advs = []
                    for i in range(len(line1_sorted) - 1):
                        advs.append((line1_sorted[i][0],
                                     round(line1_sorted[i + 1][1]
                                           - line1_sorted[i][1], 4)))
                    results[label][key] = {
                        "text": text,
                        "jc": jc,
                        "n_line1": len(line1_sorted),
                        "advances": advs,
                    }
                    print(f"\n[{label}][autoSpaceDE={key}]")
                    print(f"  text={text} n_line1={len(line1_sorted)}")
                    print(f"  advs={advs}")
                except Exception as e:
                    results[label][key] = {"error": str(e)}
                    print(f"[{label}][{key}] ERR: {e}", flush=True)
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
    existing["autospacede_interaction_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
