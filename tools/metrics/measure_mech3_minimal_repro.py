"""依頼 A: Mech 3 minimal repro — characterize jc=left yakumono compression
observed in 7f272a_p1.

If Mech 1 (FINAL RULE Type A/B/C) is always-on under kern, and Mech 2 is
jc=both+overflow, then yakumono compression observed at jc=left without
overflow must be a third mechanism (call it "Mech 3").

Hypothesis candidates for Mech 3 trigger:
- grid-snap with linesAndChars docGrid
- specific char-grid layout calc
- Some Word internal that fires when text has many yakumono between CJK
  chars on long lines

Variants (all with 36-char probe text containing `項（第N項）` patterns
across CJK chars; Mech 1 should NOT fire because all yakumono have CJK
neighbors):

  V1_kern_jc_left_grid          : kern=2, jc=left,  docGrid lines, linePitch=360
  V2_no_kern_jc_left_grid       : (control: no kern → no compression expected)
  V3_kern_jc_both_grid          : kern=2, jc=both,  docGrid lines (Mech 2 territory)
  V4_kern_jc_left_no_grid       : kern=2, jc=left,  NO docGrid (Mech 3 grid-dep test)
  V5_kern_jc_left_grid_msmincho : kern=2, jc=left,  docGrid lines (font: MS Mincho)
  V6_kern_jc_left_grid_yumincho : kern=2, jc=left,  docGrid lines (font: Yu Mincho)
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

OUT_DIR = os.path.abspath("pipeline_data/mech3_minimal_repro_docs")
RESULT_PATH = os.path.abspath(
    "pipeline_data/mech3_minimal_repro_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")
YAKUMONO_C = set("・：；！？ー―／＼")

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


def gen_styles(kern_val):
    kern = f'<w:kern w:val="{kern_val}"/>' if kern_val else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            f'{kern}'
            '<w:sz w:val="21"/>'
            '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
            '<w:name w:val="Normal"/><w:qFormat/></w:style>'
            '</w:styles>')


SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>
"""


def gen_doc(text, jc, font, with_grid, page_w_tw=11906, margin_tw=1700):
    grid_xml = ('<w:docGrid w:type="lines" w:linePitch="360"/>'
                 if with_grid else "")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>'
            '<w:r>'
            '<w:rPr>'
            f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:hint="eastAsia"/>'
            '<w:sz w:val="21"/>'
            '</w:rPr>'
            f'<w:t>{text}</w:t>'
            '</w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="{margin_tw}" w:bottom="1440" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            f'{grid_xml}'
            '</w:sectPr></w:body></w:document>')


def write_docx(path, kern_val, jc, font, with_grid, text,
                page_w_tw=11906, margin_tw=1700):
    tmp = tempfile.mkdtemp(prefix="mech3_")
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
            f.write(gen_styles(kern_val))
        with open(os.path.join(tmp, "word", "settings.xml"), "w",
                  encoding="utf-8") as f:
            f.write(SETTINGS_XML)
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(gen_doc(text, jc, font, with_grid,
                            page_w_tw, margin_tw))
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# Probe text design: `項（第Ｎ項）` pattern repeated with CJK in between.
# All yakumono are between CJK chars (no FINAL RULE trigger). If Mech 1
# fires here, that's a regression of FINAL RULE. If Mech 2 fires (and
# jc=left so no justify), that's Mech 3.
# 36 chars: 規定により項（第１項）規定により項（第２項）規定により項（第３項）
PROBE_TEXT = "規定により項（第１項）規定により項（第２項）規定により項（第３項）規定"
# 36 chars × 10.5pt = 378pt natural. Default page width ≈ 425pt → fits in 1 line w/ slack.
# To force overflow → set narrower page or longer text.

# For Mech 2 / overflow test, we need NARROWER content. Use longer text.
PROBE_TEXT_LONG = (
    "規定により項（第１項）規定により項（第２項）規定により項（第３項）"
    "規定により項（第４項）規定により項（第５項）規定により項（第６項）規定")
# 72 chars × 10.5pt = 756pt natural. Default content_w 425pt → 2 lines, line 1
# ~40 chars. With jc=both, line 1 will need slack distribution.

# Tighter content: page width 8000tw (=400pt), margins 850tw=42.5pt each
# → content_w = 400 - 85 = 315pt. 30 chars × 10.5pt = 315pt fits exactly.
# 32+ chars overflow → force compression decisions.
TIGHT_W = 8000  # ~400pt page
TIGHT_MARGIN = 850  # ~42.5pt each side → content ~315pt

VARIANTS = [
    # (label, kern_val, jc, font, with_grid, text, page_w_tw, margin_tw)
    # Wide page tests (5pt slack — minimal force)
    ("V1_kern_jc_left_grid",     2,    "left", "ＭＳ 明朝", True,  PROBE_TEXT, 11906, 1700),
    ("V2_no_kern_jc_left_grid",  None, "left", "ＭＳ 明朝", True,  PROBE_TEXT, 11906, 1700),
    ("V3_kern_jc_both_grid",     2,    "both", "ＭＳ 明朝", True,  PROBE_TEXT, 11906, 1700),
    # Tight page tests (force overflow)
    ("V4t_kern_jc_left_grid_tight",  2,    "left", "ＭＳ 明朝", True,  PROBE_TEXT, TIGHT_W, TIGHT_MARGIN),
    ("V5t_no_kern_jc_left_tight",    None, "left", "ＭＳ 明朝", True,  PROBE_TEXT, TIGHT_W, TIGHT_MARGIN),
    ("V6t_kern_jc_both_tight",       2,    "both", "ＭＳ 明朝", True,  PROBE_TEXT, TIGHT_W, TIGHT_MARGIN),
    ("V7t_kern_jc_left_no_grid_tight",2,   "left", "ＭＳ 明朝", False, PROBE_TEXT, TIGHT_W, TIGHT_MARGIN),
    ("V8t_kern_jc_left_yu_tight",    2,    "left", "Yu Mincho", True, PROBE_TEXT, TIGHT_W, TIGHT_MARGIN),
    # Even longer + tight (multi-line overflow)
    ("V9_kern_jc_both_long_tight",   2,    "both", "ＭＳ 明朝", True,  PROBE_TEXT_LONG, TIGHT_W, TIGHT_MARGIN),
    ("V10_kern_jc_left_long_tight",  2,    "left", "ＭＳ 明朝", True,  PROBE_TEXT_LONG, TIGHT_W, TIGHT_MARGIN),
]


def classify(ch):
    if ch in YAKUMONO_A:
        return "A"
    if ch in YAKUMONO_B:
        return "B"
    if ch in YAKUMONO_C:
        return "C"
    return None


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
    # Group by line via y
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
            rule_match = "none"
            if yclass == "A":
                pc = classify(prev_ch) if prev_ch else None
                if pc == "A":
                    rule_match = "A_after_A"
            elif yclass == "B":
                nc = classify(next_ch) if next_ch else None
                if nc in ("A", "B"):
                    rule_match = f"B_before_{nc}"
            advs.append({
                "ch": ch,
                "prev_ch": prev_ch,
                "next_ch": next_ch,
                "adv": adv,
                "size": sz,
                "ratio": ratio,
                "yakumono_class": yclass,
                "rule_match": rule_match,
                "compressed": (ratio is not None and ratio < 0.85
                                and yclass is not None),
            })
        line_data.append({
            "y": y,
            "n_chars": len(sorted_chars),
            "first_x": sorted_chars[0][1],
            "last_x": sorted_chars[-1][1],
            "advances": advs,
        })
    return line_data


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, kern_val, jc, font, with_grid, text, page_w, margin in VARIANTS:
        path = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(path, kern_val, jc, font, with_grid, text,
                    page_w, margin)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        lines = None
        try:
            try:
                lines = measure_doc(word, path)
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}", flush=True)
                continue
            # Summary
            yak_total = 0
            yak_compressed = []
            yak_not_compressed = []
            for ln in lines:
                for a in ln["advances"]:
                    if a["yakumono_class"]:
                        yak_total += 1
                        if a["compressed"]:
                            yak_compressed.append(a)
                        else:
                            yak_not_compressed.append(a)
            rule_match_compressed = [a for a in yak_compressed
                                       if a["rule_match"] != "none"]
            rule_match_NOT_compressed = [a for a in yak_not_compressed
                                          if a["rule_match"] != "none"]
            results[label] = {
                "kern": kern_val, "jc": jc, "font": font,
                "with_grid": with_grid, "text": text,
                "page_w_tw": page_w, "margin_tw": margin,
                "n_lines": len(lines),
                "yak_total": yak_total,
                "yak_compressed": len(yak_compressed),
                "rule_match_compressed": len(rule_match_compressed),
                "rule_match_NOT_compressed": len(rule_match_NOT_compressed),
                "lines": lines,
            }
            print(f"\n[{label}] kern={kern_val} jc={jc} font={font} "
                  f"grid={with_grid} pgW={page_w}tw mgn={margin}tw "
                  f"text_chars={len(text)}", flush=True)
            print(f"  n_lines={len(lines)} yak={yak_total} "
                  f"compressed={len(yak_compressed)} "
                  f"M1={len(rule_match_compressed)} "
                  f"M2/M3={len(yak_compressed) - len(rule_match_compressed)} "
                  f"NOT_M1_compressed_anom={len(rule_match_NOT_compressed)}",
                  flush=True)
            for a in yak_compressed:
                cls = "M1" if a["rule_match"] != "none" else "M2/M3?"
                print(f"    [{cls}] {a['ch']!r:>3} prev={a['prev_ch']!r} "
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
