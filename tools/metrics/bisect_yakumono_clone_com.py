"""§4.7 trigger bisect — clone COM-generated docx, replace ONLY document.xml.

If the modified doc still compresses, trigger is in styles.xml/settings.xml/
theme1.xml etc. (NOT in document.xml runs).
If the modified doc doesn't compress, trigger is in document.xml.
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

SRC_COM = os.path.abspath(
    "pipeline_data/yakumono_setting_docs/"
    "close_open__ＭＳ_明朝_10.5_doNotCompress.docx")
OUT_DIR = os.path.abspath("pipeline_data/yakumono_clone_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")

PROBE = "漢」（漢"

# Variants of document.xml to overlay
VARIANTS = {
    # The exact document.xml from COM (control)
    "V_original_doc": None,  # special: keep original
    # Minimal document.xml without rPr customization
    "V_min_no_rpr": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body><w:p><w:pPr/>'
        '<w:r>'
        f'<w:t>{PROBE}</w:t>'
        '</w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:type="lines" w:linePitch="360"/>'
        '</w:sectPr></w:body></w:document>'
    ),
    # Minimal document.xml with explicit ＭＳ 明朝 in run rPr (no hint)
    "V_min_explicit_font_no_hint": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body><w:p><w:pPr/>'
        '<w:r>'
        '<w:rPr>'
        '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
        '</w:rPr>'
        f'<w:t>{PROBE}</w:t>'
        '</w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:type="lines" w:linePitch="360"/>'
        '</w:sectPr></w:body></w:document>'
    ),
    # Minimal document.xml with ＭＳ 明朝 + hint=eastAsia
    "V_min_with_hint": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body><w:p><w:pPr/>'
        '<w:r>'
        '<w:rPr>'
        '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
        '</w:rPr>'
        f'<w:t>{PROBE}</w:t>'
        '</w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:type="lines" w:linePitch="360"/>'
        '</w:sectPr></w:body></w:document>'
    ),
}


def make_clone(src_docx, dst_docx, replacement_doc_xml):
    """Copy src_docx → dst_docx, replacing document.xml if given."""
    tmp = tempfile.mkdtemp(prefix="clone_doc_")
    try:
        with zipfile.ZipFile(src_docx) as z:
            z.extractall(tmp)
        if replacement_doc_xml is not None:
            with open(os.path.join(tmp, "word", "document.xml"), "w",
                      encoding="utf-8") as f:
                f.write(replacement_doc_xml)
        with zipfile.ZipFile(dst_docx, "w", zipfile.ZIP_DEFLATED) as z:
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
        for label, doc_xml in VARIANTS.items():
            path = os.path.join(OUT_DIR, f"{label}.docx")
            make_clone(SRC_COM, path, doc_xml)
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
            results[label] = advs
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
    existing["yakumono_clone_bisect_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
