"""依頼 A 続き: clone 7f272a's supporting files (styles/settings/theme/etc)
and replace ONLY document.xml with our probe text.

If compression fires here but not in fully-synthesized minimal doc:
trigger is in supporting files (styles.xml additions, settings.xml
useFELayout, theme1.xml something, etc).

If still 0 compression: trigger is in document.xml structure (rsidR
attribs, w14:paraId, namespaces, etc).
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

# Source: any real doc with compression — use a Word-COM-generated one
SRC_REAL = os.path.abspath(
    "pipeline_data/yakumono_setting_docs/"
    "close_open__ＭＳ_明朝_10.5_doNotCompress.docx")
OUT_DIR = os.path.abspath("pipeline_data/mech3_clone_docs")
RESULT_PATH = os.path.abspath(
    "pipeline_data/mech3_clone_real_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")


def classify(ch):
    if ch in YAKUMONO_A:
        return "A"
    if ch in YAKUMONO_B:
        return "B"
    return None


# Probes (CJK-flanked yakumono, no Mech 1 triggers)
PROBE_SHORT = "規定により項（第１項）規定により項（第２項）規定により項（第３項）規定"
PROBE_LONG = (
    "規定により項（第１項）規定により項（第２項）規定により項（第３項）"
    "規定により項（第４項）規定により項（第５項）規定により項（第６項）規定")


def make_minimal_document(text, jc, page_w_tw=8000, margin_tw=850):
    grid_xml = '<w:docGrid w:type="lines" w:linePitch="360"/>'
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
            '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" '
            'w:hAnsi="ＭＳ 明朝"/></w:rPr></w:pPr>'
            '<w:r>'
            '<w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" '
            'w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
            '</w:rPr>'
            f'<w:t>{text}</w:t>'
            '</w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="{margin_tw}" w:bottom="1440" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            f'{grid_xml}'
            '</w:sectPr></w:body></w:document>')


def clone_with_doc(label, text, jc, page_w_tw, margin_tw):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    tmp = tempfile.mkdtemp(prefix="clone_real_")
    try:
        with zipfile.ZipFile(SRC_REAL) as z:
            z.extractall(tmp)
        # Replace document.xml
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(make_minimal_document(text, jc, page_w_tw, margin_tw))
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
            xs.append((t,
                       float(c.Information(5)),
                       float(c.Information(6)),
                       c.Font.Size))
        except Exception:
            continue
    d.Close(SaveChanges=False)
    if not xs:
        return []
    lines_y = {}
    for ch, x, y, sz in xs:
        lines_y.setdefault(round(y, 1), []).append((ch, x, sz))
    line_data = []
    for y in sorted(lines_y.keys()):
        sorted_chars = sorted(lines_y[y], key=lambda t: t[1])
        advs = []
        for i in range(len(sorted_chars) - 1):
            ch, x, sz = sorted_chars[i]
            next_ch, next_x, _ = sorted_chars[i + 1]
            adv = round(next_x - x, 4)
            ratio = round(adv / sz, 3) if sz else None
            yclass = classify(ch)
            prev_ch = sorted_chars[i - 1][0] if i > 0 else None
            advs.append({
                "ch": ch, "prev_ch": prev_ch, "next_ch": next_ch,
                "adv": adv, "size": sz, "ratio": ratio,
                "yakumono_class": yclass,
                "compressed": (ratio is not None and ratio < 0.85
                                and yclass is not None),
            })
        line_data.append({"y": y, "n_chars": len(sorted_chars),
                           "advances": advs})
    return line_data


VARIANTS = [
    ("CL_clone_jc_left",        PROBE_SHORT, "left", 11906, 1700),
    ("CL_clone_jc_both",        PROBE_SHORT, "both", 11906, 1700),
    ("CL_clone_jc_left_tight",  PROBE_SHORT, "left", 8000, 850),
    ("CL_clone_jc_both_tight",  PROBE_SHORT, "both", 8000, 850),
    ("CL_clone_jc_left_long",   PROBE_LONG,  "left", 8000, 850),
    ("CL_clone_jc_both_long",   PROBE_LONG,  "both", 8000, 850),
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, text, jc, page_w, margin in VARIANTS:
        path = clone_with_doc(label, text, jc, page_w, margin)
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
                "text": text, "jc": jc, "page_w_tw": page_w,
                "margin_tw": margin,
                "n_lines": len(lines),
                "yak_total": yak_total,
                "yak_compressed": len(yak_compressed),
                "lines": lines,
            }
            print(f"\n[{label}] jc={jc} pgW={page_w} mgn={margin} "
                  f"text_len={len(text)}", flush=True)
            print(f"  n_lines={len(lines)} yak={yak_total} "
                  f"compressed={len(yak_compressed)}", flush=True)
            for a in yak_compressed:
                print(f"    {a['ch']!r} prev={a['prev_ch']!r} "
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
