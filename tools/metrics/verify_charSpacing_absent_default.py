"""Test default behavior when <w:characterSpacingControl> is absent
from settings.xml.

Per session 51 mech3_trigger_PINPOINTED:
  doNotCompress + kern: only Mech 1 fires
  compressPunctuation + kern: Mech 1 + Mech 2/Mech 3 fire

What about absent? 7 baseline docs have absent setting. Per
ECMA-376 §17.15.1.21, default is documented as "doNotCompress" but
worth verifying empirically.

Test: 3 variants identical except for charSpacingControl line:
  V_absent:   <w:characterSpacingControl> NOT in settings.xml
  V_doNotCompress: <w:characterSpacingControl w:val="doNotCompress"/>
  V_compressPunctuation: <w:characterSpacingControl w:val="compressPunctuation"/>

Each tested with:
- Mech 1 probe: 漢」（漢 (B→A)
- Mech 3 probe: long real text (7f272a paragraph)

If V_absent matches V_doNotCompress → default is doNotCompress
If V_absent matches V_compressPunctuation → default is compressPunctuation
If different → third behavior
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

OUT_DIR = os.path.abspath("pipeline_data/charSpacing_absent_docs")
RESULT_PATH = os.path.abspath(
    "pipeline_data/charSpacing_absent_default_2026-05-02.json")

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
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>
      <w:kern w:val="2"/>
      <w:sz w:val="21"/>
      <w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="a">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
</w:styles>
"""


def gen_settings(charSpacing_mode):
    """charSpacing_mode: 'absent' / 'doNotCompress' / 'compressPunctuation'"""
    if charSpacing_mode == "absent":
        cs_line = ""
    else:
        cs_line = (f'<w:characterSpacingControl w:val="{charSpacing_mode}"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'{cs_line}'
            '<w:compat>'
            '<w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
            '</w:compat>'
            '</w:settings>')


def gen_doc(text):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="left"/></w:pPr>'
            '<w:r>'
            '<w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" '
            'w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
            '</w:rPr>'
            f'<w:t>{text}</w:t>'
            '</w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1700" w:bottom="1440" w:left="1700"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def write_docx(path, charSpacing_mode, text):
    tmp = tempfile.mkdtemp(prefix="cs_default_")
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
            f.write(gen_settings(charSpacing_mode))
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


YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")


def classify(ch):
    if ch in YAKUMONO_A:
        return "A"
    if ch in YAKUMONO_B:
        return "B"
    return None


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
        prev_ch = xs[i - 1][0] if i > 0 else None
        next_ch = xs[i + 1][0]
        yclass = classify(ch)
        rule_match = "none"
        if yclass == "A" and classify(prev_ch) == "A":
            rule_match = "A_after_A"
        elif yclass == "B" and classify(next_ch) in ("A", "B"):
            rule_match = f"B_before_{classify(next_ch)}"
        advs.append({
            "ch": ch, "prev_ch": prev_ch, "next_ch": next_ch,
            "adv": adv, "ratio": ratio,
            "yakumono_class": yclass,
            "rule_match": rule_match,
            "compressed": (ratio is not None and ratio < 0.85
                            and yclass is not None),
        })
    return advs


PROBE_M1 = "漢」（漢"
PROBE_M3 = ("卸売市場法第６条第１項（第14条において準用する同法第６条第１項）"
            "の規定により、中央卸売市場（地方卸売市場）に係る認定事項の変更"
            "について認定を受けたいので、次のとおり関係書類を添えて申請します。")

VARIANTS = [
    # (label, charSpacing_mode, text, mech_test)
    ("A_M1_absent",   "absent",              PROBE_M1),
    ("A_M1_dnc",      "doNotCompress",       PROBE_M1),
    ("A_M1_cp",       "compressPunctuation", PROBE_M1),
    ("B_M3_absent",   "absent",              PROBE_M3),
    ("B_M3_dnc",      "doNotCompress",       PROBE_M3),
    ("B_M3_cp",       "compressPunctuation", PROBE_M3),
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, cs_mode, text in VARIANTS:
        path = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(path, cs_mode, text)
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
            yak_compressed = [a for a in advs if a["compressed"]]
            mech1_hits = [a for a in yak_compressed
                           if a["rule_match"] != "none"]
            mech_other = [a for a in yak_compressed
                           if a["rule_match"] == "none"]
            results[label] = {
                "charSpacing_mode": cs_mode, "text_chars": len(text),
                "n_compressed": len(yak_compressed),
                "n_mech1": len(mech1_hits),
                "n_other": len(mech_other),
                "advs": advs,
            }
            print(f"\n[{label}] cs={cs_mode} text_chars={len(text)}",
                  flush=True)
            print(f"  total={len(yak_compressed)} M1={len(mech1_hits)} "
                  f"other={len(mech_other)}", flush=True)
            for a in yak_compressed:
                cls = "M1" if a["rule_match"] != "none" else "M2/M3"
                print(f"    [{cls}] {a['ch']!r} prev={a['prev_ch']!r} "
                      f"next={a['next_ch']!r} adv={a['adv']} "
                      f"r={a['ratio']}", flush=True)
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
