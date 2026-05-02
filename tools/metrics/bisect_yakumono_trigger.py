"""§4.7 — Bisect what gates Mechanism 1 (Type A/B/C compression).

Surprising finding: OOXML-direct docs do NOT compress yakumono per
Type A/B/C rules. COM Documents.Add()-built docs DO. What's the trigger?

Bisect by adding properties to a minimal OOXML one at a time:
  V0: minimal — no lang, no compat15-extras, etc.
  V1: + lang ja-JP on docDefaults
  V2: + lang ja-JP on the run rPr
  V3: + compat15 extra settings
  V4: + lang on pPr
  V5: + custom Normal style with lang=ja-JP
  V6: + theme1.xml with Japanese fonts
  V7: + actual font on docDefaults (no theme refs)
  V8: combo of all above

Probe: 漢」（漢 (B→A pair, expected compression to 6pt under Mech 1)
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

OUT_DIR = os.path.abspath("pipeline_data/yakumono_trigger_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")

PROBE = "漢」（漢"

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

VARIANTS = {
    "V0_minimal": {
        "styles_dd_rpr": '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>',
        "settings_extra": "",
        "run_rpr_extra": "",
        "ppr_extra": "",
    },
    "V1_lang_eastasia_docDefaults": {
        "styles_dd_rpr": '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/>'
                          '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr>',
        "settings_extra": "",
        "run_rpr_extra": "",
        "ppr_extra": "",
    },
    "V2_lang_run": {
        "styles_dd_rpr": '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>',
        "settings_extra": "",
        "run_rpr_extra": '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>',
        "ppr_extra": "",
    },
    "V3_full_compat15": {
        "styles_dd_rpr": '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>',
        "settings_extra": (
            '<w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '<w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>'
            '</w:compat>'
        ),
        "run_rpr_extra": "",
        "ppr_extra": "",
    },
    "V4_lang_on_ppr": {
        "styles_dd_rpr": '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>',
        "settings_extra": "",
        "run_rpr_extra": "",
        "ppr_extra": '<w:rPr><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr>',
    },
    "V5_combined_lang": {
        "styles_dd_rpr": '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/>'
                          '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr>',
        "settings_extra": "",
        "run_rpr_extra": '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>',
        "ppr_extra": '<w:rPr><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr>',
    },
    "V6_combined_lang_compat": {
        "styles_dd_rpr": '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/>'
                          '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr>',
        "settings_extra": (
            '<w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '</w:compat>'
        ),
        "run_rpr_extra": '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>',
        "ppr_extra": '<w:rPr><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr>',
    },
}


def gen_styles(dd_rpr):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            f'<w:rPrDefault>{dd_rpr}</w:rPrDefault>'
            '<w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>')


def gen_settings(extra):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="doNotCompress"/>'
            f'{extra}</w:settings>')


def gen_doc(text, ppr_extra, run_rpr_extra):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="left"/>{ppr_extra}</w:pPr>'
            '<w:r>'
            '<w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
            '<w:sz w:val="24"/>'
            f'{run_rpr_extra}'
            '</w:rPr>'
            f'<w:t>{text}</w:t>'
            '</w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1700" w:bottom="1440" w:left="1700"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>')


def write_docx(path, dd_rpr, settings_extra, ppr_extra, run_rpr_extra):
    tmp = tempfile.mkdtemp(prefix="trig_doc_")
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
            f.write(gen_styles(dd_rpr))
        with open(os.path.join(tmp, "word", "settings.xml"), "w",
                  encoding="utf-8") as f:
            f.write(gen_settings(settings_extra))
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(gen_doc(PROBE, ppr_extra, run_rpr_extra))
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for label, v in VARIANTS.items():
            path = os.path.join(OUT_DIR, f"{label}.docx")
            write_docx(path,
                       v["styles_dd_rpr"],
                       v["settings_extra"],
                       v["ppr_extra"],
                       v["run_rpr_extra"])
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
                        xs.append((t, float(c.Information(5))))
                    except Exception:
                        continue
                d.Close(SaveChanges=False)
                advs = [(xs[i][0],
                         round(xs[i + 1][1] - xs[i][1], 4))
                        for i in range(len(xs) - 1)]
            except Exception as e:
                advs = {"error": str(e)}
            results[label] = {"text": PROBE, "advances": advs}
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
    existing["yakumono_trigger_bisect_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
