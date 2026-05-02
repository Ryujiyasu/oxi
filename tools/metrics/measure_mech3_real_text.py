"""依頼 A 続き 2: probe with 7f272a's ACTUAL paragraph text in clone doc.

7f272a paragraph 12 was the R17 big_loser where Word DOES compress 6
yakumono (Mech 2 partial 8.0/9.5pt). If we put THIS exact text into a
clone of 7f272a's supporting files, does compression carry over?

This isolates whether the trigger is:
  - text content (yes if compression fires here)
  - paragraph properties / run rPr (no if it fires here, must be in supporting files)
  - both
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

# Use 7f272a's own supporting files
SRC_REAL = os.path.abspath(
    "tools/golden-test/documents/docx/7f272a2dfd3b_index-21.docx")
OUT_DIR = os.path.abspath("pipeline_data/mech3_real_text_docs")
RESULT_PATH = os.path.abspath(
    "pipeline_data/mech3_real_text_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")


def classify(ch):
    if ch in YAKUMONO_A:
        return "A"
    if ch in YAKUMONO_B:
        return "B"
    return None


# 7f272a paragraph 13 text (the actual R17 big_loser)
ACTUAL_TEXT = (
    "卸売市場法第６条第１項（第14条において準用する同法第６条第１項）"
    "の規定により、中央卸売市場（地方卸売市場）に係る認定事項の変更について"
    "認定を受けたいので、次のとおり関係書類を添えて申請します。")


def make_doc(text, jc):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"'
            ' xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
            ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            ' xmlns:o="urn:schemas-microsoft-com:office:office"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
            ' xmlns:v="urn:schemas-microsoft-com:vml"'
            ' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"'
            ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
            ' xmlns:w10="urn:schemas-microsoft-com:office:word"'
            ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
            ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
            ' mc:Ignorable="w14 w15">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>'
            '<w:r>'
            '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            f'<w:t>{text}</w:t>'
            '</w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1700" w:bottom="1440" w:left="1700"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def clone_with_text(label, text, jc):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    tmp = tempfile.mkdtemp(prefix="real_text_")
    try:
        with zipfile.ZipFile(SRC_REAL) as z:
            z.extractall(tmp)
        # Replace document.xml ONLY (keep all supporting files from real doc)
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(make_doc(text, jc))
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
        return out_path
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def measure_doc(word, path):
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
                       float(c.Information(6)), c.Font.Size,
                       c.Font.Name))
        except Exception:
            continue
    d.Close(SaveChanges=False)
    if not xs:
        return []
    lines_y = {}
    for ch, x, y, sz, fn in xs:
        lines_y.setdefault(round(y, 1), []).append((ch, x, sz, fn))
    line_data = []
    for y in sorted(lines_y.keys()):
        sorted_chars = sorted(lines_y[y], key=lambda t: t[1])
        advs = []
        for i in range(len(sorted_chars) - 1):
            ch, x, sz, fn = sorted_chars[i]
            next_ch = sorted_chars[i + 1][0]
            next_x = sorted_chars[i + 1][1]
            adv = round(next_x - x, 4)
            ratio = round(adv / sz, 3) if sz else None
            yclass = classify(ch)
            advs.append({
                "ch": ch, "next_ch": next_ch,
                "adv": adv, "size": sz, "ratio": ratio,
                "yakumono_class": yclass,
                "compressed": (ratio is not None and ratio < 0.85
                                and yclass is not None),
            })
        line_data.append({"y": y, "n_chars": len(sorted_chars),
                           "first_x": sorted_chars[0][1],
                           "last_x": sorted_chars[-1][1],
                           "advances": advs})
    return line_data


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    # Test with original 7f272a (control)
    # vs clone with replaced document.xml using same text
    for label, text, jc in [
        ("ACTUAL_jc_both", ACTUAL_TEXT, "both"),
        ("ACTUAL_jc_left", ACTUAL_TEXT, "left"),
    ]:
        path = clone_with_text(label, text, jc)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            try:
                lines = measure_doc(word, path)
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}", flush=True)
                continue
            yak_total = 0
            yak_compressed = []
            for ln in lines:
                for a in ln["advances"]:
                    if a["yakumono_class"]:
                        yak_total += 1
                        if a["compressed"]:
                            yak_compressed.append(a)
            results[label] = {
                "text": text, "jc": jc,
                "n_lines": len(lines),
                "yak_total": yak_total,
                "yak_compressed": len(yak_compressed),
                "lines": lines,
            }
            print(f"\n[{label}] jc={jc} text_len={len(text)}", flush=True)
            print(f"  n_lines={len(lines)} yak={yak_total} "
                  f"compressed={len(yak_compressed)}", flush=True)
            for a in yak_compressed:
                print(f"    {a['ch']!r} next={a['next_ch']!r} "
                      f"adv={a['adv']} r={a['ratio']}", flush=True)
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
