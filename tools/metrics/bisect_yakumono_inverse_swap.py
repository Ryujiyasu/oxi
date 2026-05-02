"""§4.7 trigger bisect — inverse swap.

Start from COM doc (which compresses), replace one file at a time with
my minimal V8_everything version. Note when compression stops.
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
OUT_DIR = os.path.abspath("pipeline_data/yakumono_inv_swap_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")

# Each variant: which files to OVERLAY from SRC_MIN onto SRC_COM clone.
# Goal: find which files in COM are essential.
VARIANTS = [
    ("V_keep_com_all", []),                     # baseline: all from COM
    ("V_replace_settings", ["word/settings.xml"]),
    ("V_replace_styles", ["word/styles.xml"]),
    ("V_replace_theme", ["word/theme/theme1.xml"]),
    ("V_replace_fontTable", ["word/fontTable.xml"]),
    ("V_replace_webSettings", ["word/webSettings.xml"]),
    # Settings.xml is the most likely trigger location due to useFELayout etc.
]


def build(label, replacement_files):
    tmp = tempfile.mkdtemp(prefix="inv_swap_")
    try:
        with zipfile.ZipFile(SRC_COM) as z:
            z.extractall(tmp)
        with zipfile.ZipFile(SRC_MIN) as zm:
            zm_namelist = zm.namelist()
            for f in replacement_files:
                target = os.path.join(tmp, *f.split("/"))
                if f in zm_namelist:
                    os.makedirs(os.path.dirname(target), exist_ok=True)
                    with open(target, "wb") as out:
                        out.write(zm.read(f))
                else:
                    # SRC_MIN doesn't have this file; remove from clone
                    if os.path.exists(target):
                        os.remove(target)
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
        for label, replacements in VARIANTS:
            path = build(label, replacements)
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
            # Compression detection: 」 advance < 8pt = compressed
            if isinstance(advs, list) and len(advs) >= 2:
                kakko_adv = advs[1][1]
                marker = ("COMPRESSED" if kakko_adv < 8 else "FULL")
            else:
                marker = "ERR"
            results[label] = {"replaced": replacements, "advances": advs,
                               "marker": marker}
            print(f"[{label}] replaced={replacements} → {marker} {advs}",
                  flush=True)
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
    existing["yakumono_inv_swap_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
