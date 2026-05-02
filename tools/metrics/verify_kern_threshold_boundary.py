"""kern val threshold boundary verification.

Per ECMA-376 §17.3.2.18: w:kern val=N is the minimum font size in
HALF-POINTS for which kerning (and hence yakumono compression) applies.
Earlier session 51 finding (measure_kern_value_and_override.py):
  val=2 → fires for 10.5pt (21 hp ≥ 2)
  val=100 → does NOT fire for 10.5pt (21 hp < 100)

Exact boundary: with text 10.5pt = 21 hp, the boundary is val=21 vs val=22:
  val=21 (threshold 10.5pt) → 21 >= 21 → SHOULD fire
  val=22 (threshold 11pt)   → 21 <  22 → SHOULD NOT fire

Probe: 漢」（漢 (B→A FINAL RULE trigger). Mech 1 fires if kern active.

Variants:
  V_val_2:   sz=21 (10.5pt) text + kern val=2  (threshold 1pt) → fire
  V_val_20:  sz=21 + kern val=20 (threshold 10pt) → fire
  V_val_21:  sz=21 + kern val=21 (threshold 10.5pt) → boundary, fire?
  V_val_22:  sz=21 + kern val=22 (threshold 11pt) → NO fire
  V_val_30:  sz=21 + kern val=30 → NO fire
  V_val_2_sz_24: sz=24 (12pt) + kern val=2  → fire (control)
  V_val_22_sz_24: sz=24 + kern val=22 (24 >= 22) → fire
"""
import os
import time
import json
import zipfile
import shutil
import sys
import tempfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/kern_threshold_boundary_docs")
RESULT_PATH = os.path.abspath(
    "pipeline_data/kern_threshold_boundary_2026-05-02.json")

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

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
  </w:compat>
</w:settings>
"""


def gen_styles(kern_val, sz_val):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
            f'<w:kern w:val="{kern_val}"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
            '<w:name w:val="Normal"/><w:qFormat/></w:style>'
            '</w:styles>')


def gen_doc(text):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="left"/></w:pPr>'
            '<w:r>'
            '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            f'<w:t>{text}</w:t>'
            '</w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1700" w:bottom="1440" w:left="1700"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>')


def write_docx(path, kern_val, sz_val, text):
    tmp = tempfile.mkdtemp(prefix="kern_thresh_")
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
            f.write(gen_styles(kern_val, sz_val))
        with open(os.path.join(tmp, "word", "settings.xml"), "w",
                  encoding="utf-8") as f:
            f.write(SETTINGS_XML)
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(gen_doc(text))
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")


def measure(word, path):
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
            xs.append((t, float(c.Information(5)), c.Font.Size))
        except Exception:
            continue
    d.Close(SaveChanges=False)
    advs = []
    for i in range(len(xs) - 1):
        ch, x, sz = xs[i]
        adv = round(xs[i + 1][1] - x, 4)
        ratio = round(adv / sz, 3) if sz else None
        advs.append({"ch": ch, "adv": adv, "ratio": ratio,
                      "compressed": ratio is not None and ratio < 0.85
                                     and ch in YAKUMONO_B})
    return advs


VARIANTS = [
    # (label, kern_val, sz_val, text, expected_fire?)
    ("V_val_2_sz_21",   2,   21, "漢」（漢", True),
    ("V_val_20_sz_21",  20,  21, "漢」（漢", True),
    ("V_val_21_sz_21",  21,  21, "漢」（漢", True),  # boundary equal
    ("V_val_22_sz_21",  22,  21, "漢」（漢", False), # boundary fail
    ("V_val_30_sz_21",  30,  21, "漢」（漢", False),
    ("V_val_100_sz_21", 100, 21, "漢」（漢", False),
    ("V_val_22_sz_24",  22,  24, "漢」（漢", True),  # 24 >= 22
    ("V_val_24_sz_24",  24,  24, "漢」（漢", True),  # equal
    ("V_val_25_sz_24",  25,  24, "漢」（漢", False), # 24 < 25
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, kern_val, sz_val, text, expected_fire in VARIANTS:
        path = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(path, kern_val, sz_val, text)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            try:
                advs = measure(word, path)
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}", flush=True)
                continue
            comp = [a for a in advs if a["compressed"]]
            n_comp = len(comp)
            actual_fire = n_comp > 0
            verdict = "✓" if actual_fire == expected_fire else "✗ FAIL"
            results[label] = {
                "kern_val": kern_val, "sz_val": sz_val,
                "expected_fire": expected_fire,
                "actual_fire": actual_fire,
                "n_compressed": n_comp,
                "verdict": verdict,
                "advs": advs,
            }
            print(f"[{label}] kern={kern_val} sz={sz_val} "
                  f"(threshold={kern_val/2:.1f}pt, font={sz_val/2:.1f}pt) "
                  f"expected={expected_fire} actual={actual_fire} "
                  f"({n_comp} compressed) {verdict}", flush=True)
            for a in comp:
                print(f"    {a['ch']!r} adv={a['adv']} r={a['ratio']}",
                      flush=True)
        finally:
            try:
                word.Quit()
            except Exception:
                pass
            time.sleep(1.0)

    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}", flush=True)


if __name__ == "__main__":
    main()
