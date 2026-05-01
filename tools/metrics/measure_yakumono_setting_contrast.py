"""§4.7 Yakumono compression — measure under both compressPunctuation and
doNotCompress settings on the same adjacency patterns to resolve PROVISIONAL
status (spec line 588-598).

For each (font, size, setting, probe): build a docx, patch settings.xml,
re-open via Word COM, measure per-char advance via Information(5).

Output: pipeline_data/ra_manual_measurements.json (key
"yakumono_setting_contrast_2026-05-02").

Single Word instance reused across all probes for speed.
"""
import win32com.client
import os
import time
import json
import zipfile
import re
import shutil
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PROBES = [
    ("close_open",       "漢」（漢"),
    ("paren_pair_inner", "漢「」漢"),
    ("punct_excl",       "漢、！漢"),
    ("close_punct",      "漢）。漢"),
    ("punct_open",       "漢、（漢"),
    ("close_punct2",     "漢」、漢"),
    ("punct_punct",      "漢、。漢"),
    ("AAAA",             "（（（（"),
    ("BBBB",             "））））"),
    ("c_only",           "漢」漢"),
    ("o_only",           "漢（漢"),
    ("dash_em",          "漢——漢"),
    ("dash_bar",         "漢――漢"),
    ("excl_only",        "漢！漢"),
]

import argparse

FONTS_SIZES = [
    ("ＭＳ 明朝",     10.5),
    ("ＭＳ 明朝",     14.0),
    ("ＭＳ ゴシック", 10.5),
    ("Yu Mincho",    10.5),
]

SETTINGS = ["doNotCompress", "compressPunctuation"]

OUT_DIR = os.path.abspath("pipeline_data/yakumono_setting_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")


def patch_setting(docx_path: str, setting_value: str):
    tmp_dir = tempfile.mkdtemp(prefix="yakumono_patch_")
    try:
        with zipfile.ZipFile(docx_path) as z:
            z.extractall(tmp_dir)
        settings_path = os.path.join(tmp_dir, "word", "settings.xml")
        with open(settings_path, encoding="utf-8") as f:
            s = f.read()
        new_s = re.sub(
            r'<w:characterSpacingControl[^/]*/>',
            f'<w:characterSpacingControl w:val="{setting_value}"/>',
            s,
        )
        if new_s == s:
            new_s = s.replace(
                "</w:settings>",
                f'<w:characterSpacingControl w:val="{setting_value}"/></w:settings>',
            )
        with open(settings_path, "w", encoding="utf-8") as f:
            f.write(new_s)
        os.remove(docx_path)
        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp_dir):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp_dir).replace("\\", "/")
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def make_word():
    w = win32com.client.Dispatch("Word.Application")
    w.Visible = False
    w.DisplayAlerts = False
    return w


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--skip-completed", action="store_true",
                    help="Skip font/size/setting combos already in results JSON")
    args = ap.parse_args()
    os.makedirs(OUT_DIR, exist_ok=True)
    # Load existing to merge, optionally skip completed
    if os.path.exists(RESULT_PATH):
        try:
            with open(RESULT_PATH, encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = {}
    else:
        existing = {}
    existing_block = existing.get("yakumono_setting_contrast_2026-05-02", {})
    results = dict(existing_block)
    for font, size in FONTS_SIZES:
        font_key = f"{font}_{size}"
        if font_key not in results:
            results[font_key] = {}
        for setting in SETTINGS:
            if (args.skip_completed and setting in results[font_key]
                    and len(results[font_key][setting]) >= len(PROBES)):
                # treat as complete only if no entry has error dict
                def has_err(v):
                    a = v.get("advances")
                    return isinstance(a, dict) and "error" in a
                bad = any(has_err(v)
                          for v in results[font_key][setting].values())
                if not bad:
                    print(f"Skip [{font_key}][{setting}] (already complete)",
                          flush=True)
                    continue
            word = make_word()
            if setting not in results[font_key]:
                results[font_key][setting] = {}
            try:
                for label, text in PROBES:
                    fname = (f"{label}__{font.replace(' ', '_')}_"
                             f"{size}_{setting}.docx")
                    doc_path = os.path.join(OUT_DIR, fname)
                    try:
                        d = word.Documents.Add()
                        time.sleep(0.10)
                        rng = d.Range()
                        rng.InsertAfter(text)
                        rng = d.Range()
                        rng.Font.Name = font
                        rng.Font.Size = size
                        d.Paragraphs(1).Alignment = 0
                        d.SaveAs2(doc_path, FileFormat=12)
                        d.Close(SaveChanges=False)
                        patch_setting(doc_path, setting)
                        d = word.Documents.Open(doc_path, ReadOnly=True)
                        time.sleep(0.15)
                        chars = d.Range().Characters
                        xs = []
                        for ci in range(1, chars.Count + 1):
                            try:
                                c = chars(ci)
                                ch = c.Text
                                if ch in ("\r", "\x07"):
                                    continue
                                xs.append((ch, float(c.Information(5))))
                            except Exception:
                                continue
                        d.Close(SaveChanges=False)
                        advs = [(xs[i][0],
                                 round(xs[i + 1][1] - xs[i][1], 4))
                                for i in range(len(xs) - 1)]
                    except Exception as e:
                        advs = {"error": str(e)}
                    results[font_key][setting][label] = {
                        "text": text,
                        "advances": advs,
                    }
                    line = (f"[{font_key}][{setting:18s}][{label:18s}] "
                            f"{text}: {advs}")
                    print(line, flush=True)
            finally:
                try:
                    word.Quit()
                except Exception:
                    pass
                time.sleep(1.0)

    existing["yakumono_setting_contrast_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
