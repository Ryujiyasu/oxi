"""Verify: synthesized minimal OOXML with `compressPunctuation` SHOULD
trigger Mech 3 (per bisection finding). Earlier session 51 minimal repros
all used `doNotCompress` — that's why they showed 0 compression.

Build minimal OOXML with `compressPunctuation` + 7f272a's actual paragraph
text. Should compress.
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

OUT_DIR = os.path.abspath("pipeline_data/mech3_synth_compressPunctuation_docs")
RESULT_PATH = os.path.abspath(
    "pipeline_data/mech3_synth_compressPunctuation_2026-05-02.json")

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


def gen_settings(charSpacing):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:characterSpacingControl w:val="{charSpacing}"/>'
            '<w:compat>'
            '<w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
            '</w:compat>'
            '</w:settings>')


# 7f272a's actual paragraph 13 text
ACTUAL_TEXT = (
    "卸売市場法第６条第１項（第14条において準用する同法第６条第１項）"
    "の規定により、中央卸売市場（地方卸売市場）に係る認定事項の変更について"
    "認定を受けたいので、次のとおり関係書類を添えて申請します。")


def gen_doc(text, jc):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>'
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


def write_docx(path, charSpacing, jc):
    tmp = tempfile.mkdtemp(prefix="synth_cp_")
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
            f.write(gen_settings(charSpacing))
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(gen_doc(ACTUAL_TEXT, jc))
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


VARIANTS = [
    ("synth_compressPunctuation_jc_left", "compressPunctuation", "left"),
    ("synth_compressPunctuation_jc_both", "compressPunctuation", "both"),
    ("synth_doNotCompress_jc_left",       "doNotCompress",       "left"),
    ("synth_doNotCompress_jc_both",       "doNotCompress",       "both"),
]


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
            xs.append((t, float(c.Information(5)),
                       float(c.Information(6)), c.Font.Size))
        except Exception:
            continue
    d.Close(SaveChanges=False)
    if not xs:
        return {"compressed": []}
    lines = {}
    for ch, x, y, sz in xs:
        lines.setdefault(round(y, 1), []).append((ch, x, sz))
    compressed = []
    for y in sorted(lines.keys()):
        sc = sorted(lines[y], key=lambda t: t[1])
        for i in range(len(sc) - 1):
            ch, x, sz = sc[i]
            adv = round(sc[i + 1][1] - x, 4)
            ratio = round(adv / sz, 3) if sz else None
            yclass = classify(ch)
            if yclass and ratio is not None and ratio < 0.85:
                compressed.append({
                    "ch": ch, "next_ch": sc[i + 1][0],
                    "adv": adv, "ratio": ratio,
                })
    return {"compressed": compressed}


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, charSpacing, jc in VARIANTS:
        path = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(path, charSpacing, jc)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            try:
                res = measure(word, path)
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}", flush=True)
                continue
            n = len(res["compressed"])
            results[label] = {"charSpacing": charSpacing, "jc": jc, **res,
                              "n_compressed": n}
            print(f"\n[{label}] charSpacing={charSpacing} jc={jc}: "
                  f"compressed={n}", flush=True)
            for c in res["compressed"]:
                print(f"  {c['ch']!r} next={c['next_ch']!r} "
                      f"adv={c['adv']} r={c['ratio']}", flush=True)
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
