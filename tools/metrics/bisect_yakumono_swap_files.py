"""§4.7 trigger bisect — start from minimal doc, progressively swap in
COM doc's files, find which one is the trigger.

Starting V8_resaved (no compression) → swap files from COM source one
at a time until compression fires.
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

SRC_MIN = os.path.abspath(
    "pipeline_data/yakumono_trigger_v2_docs/V8_everything.docx")
SRC_COM = os.path.abspath(
    "pipeline_data/yakumono_setting_docs/"
    "close_open__ＭＳ_明朝_10.5_doNotCompress.docx")
OUT_DIR = os.path.abspath("pipeline_data/yakumono_swap_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")

VARIANTS = [
    # (label, list of files to copy from SRC_COM into SRC_MIN)
    ("V_baseline_min", []),
    ("V_swap_settings", ["word/settings.xml"]),
    ("V_swap_styles", ["word/styles.xml"]),
    ("V_swap_theme", ["word/theme/theme1.xml"]),
    ("V_swap_fontTable", ["word/fontTable.xml"]),
    ("V_swap_webSettings", ["word/webSettings.xml"]),
    ("V_swap_settings_styles", ["word/settings.xml", "word/styles.xml"]),
    ("V_swap_all_word_xml", [
        "word/settings.xml", "word/styles.xml",
        "word/theme/theme1.xml", "word/fontTable.xml",
        "word/webSettings.xml",
    ]),
]


def build(label, swap_files):
    tmp = tempfile.mkdtemp(prefix="swap_doc_")
    try:
        with zipfile.ZipFile(SRC_MIN) as z:
            z.extractall(tmp)
        with zipfile.ZipFile(SRC_COM) as zc:
            for f in swap_files:
                target = os.path.join(tmp, *f.split("/"))
                os.makedirs(os.path.dirname(target), exist_ok=True)
                with open(target, "wb") as out:
                    out.write(zc.read(f))
        # Need to update [Content_Types].xml relations and word/_rels too?
        # For settings/styles/theme/fontTable/webSettings — the _rels and
        # [Content_Types] from SRC_MIN might not include them.
        # Let me also copy those over from SRC_COM if any swap is requested.
        if swap_files:
            for f in ["[Content_Types].xml", "word/_rels/document.xml.rels"]:
                target = os.path.join(tmp, *f.split("/"))
                os.makedirs(os.path.dirname(target), exist_ok=True)
                with zipfile.ZipFile(SRC_COM) as zc:
                    try:
                        with open(target, "wb") as out:
                            out.write(zc.read(f))
                    except KeyError:
                        pass
        out_path = os.path.join(OUT_DIR, f"{label}.docx")
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
        return out_path
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for label, swap_files in VARIANTS:
            path = build(label, swap_files)
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
            results[label] = {"swapped": swap_files, "advances": advs}
            print(f"[{label}] swap={swap_files} → {advs}", flush=True)
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
    existing["yakumono_swap_bisect_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
