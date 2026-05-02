"""§4.7 — Bisect yakumono trigger V2: focus on hint=eastAsia and docGrid.

Found that COM-generated docs have:
1. w:hint="eastAsia" on run rPr (NOT in my prior OOXML)
2. <w:docGrid w:type="lines" w:linePitch="360"/> in sectPr

Test these two as potential triggers for Mech 1 compression.
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

OUT_DIR = os.path.abspath("pipeline_data/yakumono_trigger_v2_docs")
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
def gen_settings(use_fe_layout, theme_font_lang_ja, balance_byte_width):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="doNotCompress"/>'
            '<w:compat>'
            f'{"<w:useFELayout/>" if use_fe_layout else ""}'
            f'{"<w:balanceSingleByteDoubleByteWidth/>" if balance_byte_width else ""}'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '</w:compat>'
            f'{("<w:themeFontLang w:val=\"en-US\" w:eastAsia=\"ja-JP\"/>") if theme_font_lang_ja else ""}'
            '</w:settings>')


VARIANTS = {
    "V0_baseline":              {"hint": False, "grid": False, "fe": False, "tfl": False, "bbw": False},
    "V1_use_fe_layout_only":    {"hint": False, "grid": False, "fe": True,  "tfl": False, "bbw": False},
    "V2_themeFontLang_only":    {"hint": False, "grid": False, "fe": False, "tfl": True,  "bbw": False},
    "V3_balance_byte_only":     {"hint": False, "grid": False, "fe": False, "tfl": False, "bbw": True},
    "V4_fe_plus_tfl":           {"hint": False, "grid": False, "fe": True,  "tfl": True,  "bbw": False},
    "V5_all_compat":            {"hint": False, "grid": False, "fe": True,  "tfl": True,  "bbw": True},
    "V6_all_compat_plus_hint":  {"hint": True,  "grid": False, "fe": True,  "tfl": True,  "bbw": True},
    "V7_all_compat_plus_grid":  {"hint": False, "grid": True,  "fe": True,  "tfl": True,  "bbw": True},
    "V8_everything":            {"hint": True,  "grid": True,  "fe": True,  "tfl": True,  "bbw": True},
}


def gen_doc(text, hint, grid):
    hint_attr = ' w:hint="eastAsia"' if hint else ""
    grid_xml = ('<w:docGrid w:type="lines" w:linePitch="360"/>'
                if grid else "")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="left"/></w:pPr>'
            '<w:r>'
            '<w:rPr>'
            f'<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"{hint_attr}/>'
            '<w:sz w:val="24"/>'
            '</w:rPr>'
            f'<w:t>{text}</w:t>'
            '</w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1700" w:bottom="1440" w:left="1700"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            f'{grid_xml}'
            '</w:sectPr></w:body></w:document>')


def write_docx(path, hint, grid, fe, tfl, bbw):
    tmp = tempfile.mkdtemp(prefix="t2_doc_")
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
            f.write(gen_settings(fe, tfl, bbw))
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(gen_doc(PROBE, hint, grid))
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
            write_docx(path, v["hint"], v["grid"], v["fe"], v["tfl"],
                       v["bbw"])
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
    existing["yakumono_trigger_v2_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
