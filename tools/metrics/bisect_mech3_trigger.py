"""依頼 A: Mech 3 trigger bisection.

Start with 7f272a as baseline (compression fires on paragraph 13).
Progressively modify supporting files. Variant where compression
disappears = identifies the Mech 3 trigger element.

Test target: paragraph index 13 (the R17 big_loser known to compress
6 yakumono via Mech 2 partial 8.0/9.5pt).
"""
import os
import time
import json
import zipfile
import shutil
import sys
import tempfile
import re
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC_DOC = os.path.abspath(
    "tools/golden-test/documents/docx/7f272a2dfd3b_index-21.docx")
OUT_DIR = os.path.abspath("pipeline_data/mech3_bisect_docs")
RESULT_PATH = os.path.abspath(
    "pipeline_data/mech3_trigger_bisect_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")


def classify(ch):
    if ch in YAKUMONO_A:
        return "A"
    if ch in YAKUMONO_B:
        return "B"
    return None


def make_variant(label, modifications):
    """modifications is dict of file_path -> replacement_content (str/bytes)
    or file_path -> None to delete the file."""
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    tmp = tempfile.mkdtemp(prefix="bisect_")
    try:
        with zipfile.ZipFile(SRC_DOC) as z:
            z.extractall(tmp)
        for fp, content in modifications.items():
            target = os.path.join(tmp, *fp.split("/"))
            if content is None:
                if os.path.exists(target):
                    os.remove(target)
            else:
                os.makedirs(os.path.dirname(target), exist_ok=True)
                if isinstance(content, str):
                    with open(target, "w", encoding="utf-8") as f:
                        f.write(content)
                else:
                    with open(target, "wb") as f:
                        f.write(content)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
        return out_path
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def measure_paragraph_13(word, path):
    """Open doc, find paragraph 13, return per-char advances + compression count."""
    d = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.5)
    try:
        p = d.Paragraphs(13)
        rng = p.Range
        text = rng.Text.strip()
        chars = rng.Characters
        per_char = []
        for ci in range(1, chars.Count + 1):
            try:
                c = chars(ci)
                t = c.Text
                if t in ("\r", "\x07"):
                    continue
                per_char.append({
                    "i": ci,
                    "ch": t,
                    "x": round(float(c.Information(5)), 4),
                    "y": round(float(c.Information(6)), 4),
                    "size": c.Font.Size,
                })
            except Exception:
                continue
    finally:
        try:
            d.Close(SaveChanges=False)
        except Exception:
            pass
    if not per_char:
        return {"text": "", "n_compressed": 0, "compressed_chars": []}
    # Group by line
    lines = {}
    for r in per_char:
        lines.setdefault(round(r["y"], 1), []).append(r)
    compressed = []
    for y in sorted(lines.keys()):
        sc = sorted(lines[y], key=lambda r: r["x"])
        for i in range(len(sc) - 1):
            ch = sc[i]["ch"]
            adv = round(sc[i + 1]["x"] - sc[i]["x"], 4)
            sz = sc[i]["size"]
            ratio = round(adv / sz, 3) if sz else None
            yclass = classify(ch)
            if yclass and ratio is not None and ratio < 0.85:
                next_ch = sc[i + 1]["ch"]
                prev_ch = sc[i - 1]["ch"] if i > 0 else None
                compressed.append({
                    "i": sc[i]["i"], "ch": ch,
                    "prev_ch": prev_ch, "next_ch": next_ch,
                    "adv": adv, "ratio": ratio,
                })
    return {"text": text[:80], "n_compressed": len(compressed),
            "compressed_chars": compressed}


# Define minimal settings.xml replacement (strips most elements)
MINIMAL_SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="840"/>
  <w:characterSpacingControl w:val="compressPunctuation"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
  </w:compat>
</w:settings>
"""


# Read original settings to apply targeted modifications
def get_original_settings():
    with zipfile.ZipFile(SRC_DOC) as z:
        return z.read("word/settings.xml").decode("utf-8")


ORIG_SETTINGS = get_original_settings()


def settings_remove(elem_pattern):
    """Remove element from original settings.xml."""
    return re.sub(elem_pattern, "", ORIG_SETTINGS)


def settings_replace_charSpacing(new_val):
    return re.sub(r'<w:characterSpacingControl[^/]*/>',
                   f'<w:characterSpacingControl w:val="{new_val}"/>',
                   ORIG_SETTINGS)


def settings_replace_compatibilityMode(new_val):
    return re.sub(
        r'<w:compatSetting w:name="compatibilityMode"[^/]*/>',
        f'<w:compatSetting w:name="compatibilityMode" '
        f'w:uri="http://schemas.microsoft.com/office/word" '
        f'w:val="{new_val}"/>',
        ORIG_SETTINGS)


# Define variants
VARIANTS = [
    ("V0_baseline", {}),  # original 7f272a (control)
    # File-level removals
    ("V1_no_fontTable",   {"word/fontTable.xml": None}),
    ("V2_no_theme",       {"word/theme/theme1.xml": None}),
    ("V3_no_webSettings", {"word/webSettings.xml": None}),
    ("V4_no_endnotes",    {"word/endnotes.xml": None,
                            "word/footnotes.xml": None}),
    # Settings.xml modifications
    ("V5_settings_minimal",
     {"word/settings.xml": MINIMAL_SETTINGS}),
    ("V6_remove_useFELayout",
     {"word/settings.xml": settings_remove(r'<w:useFELayout/>')}),
    ("V7_remove_balanceByteWidth",
     {"word/settings.xml": settings_remove(
         r'<w:balanceSingleByteDoubleByteWidth/>')}),
    ("V8_remove_themeFontLang",
     {"word/settings.xml": settings_remove(
         r'<w:themeFontLang[^/]*/>')}),
    ("V9_charSpacing_doNotCompress",
     {"word/settings.xml": settings_replace_charSpacing("doNotCompress")}),
    ("V10_compat_15",
     {"word/settings.xml": settings_replace_compatibilityMode("15")}),
    ("V11_remove_adjustLineHeight",
     {"word/settings.xml": settings_remove(
         r'<w:adjustLineHeightInTable/>')}),
    ("V12_remove_spaceForUL",
     {"word/settings.xml": settings_remove(r'<w:spaceForUL/>')}),
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, mods in VARIANTS:
        path = make_variant(label, mods)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            try:
                res = measure_paragraph_13(word, path)
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}", flush=True)
                continue
            results[label] = {**res, "modifications": list(mods.keys())}
            print(f"\n[{label}] mods={list(mods.keys())}", flush=True)
            print(f"  text: {res['text']!r}", flush=True)
            print(f"  n_compressed: {res['n_compressed']}", flush=True)
            for c in res["compressed_chars"][:10]:
                print(f"    [{c['i']:3d}] {c['ch']!r} "
                      f"prev={c['prev_ch']!r} next={c['next_ch']!r} "
                      f"adv={c['adv']} r={c['ratio']}", flush=True)
        finally:
            try:
                word.Quit()
            except Exception:
                pass
            time.sleep(1.5)
        # Save after each variant for resilience
        with open(RESULT_PATH, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}", flush=True)


if __name__ == "__main__":
    main()
