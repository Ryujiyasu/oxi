"""Test if Mechanism 2 (justify-time slack distribution) is also gated
by `<w:kern>`.

Probe: text where Mech 1 (Type A/B/C adjacency) does NOT fire, so any
compression observed is purely Mech 2.

Mech 1 doesn't fire when yakumono's neighbors are both CJK (NOT yakumono).
So `漢、漢、漢、漢、漢` has 4 `、` each between CJK = 0 Mech 1 triggers.

Forcing overflow at jc=both → if any `、` compresses, Mech 2 fired.
Test kern on vs off.
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
OUT_DIR = os.path.abspath("pipeline_data/mech2_kern_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")


def make_styles(with_kern=True):
    kern = '<w:kern w:val="2"/>' if with_kern else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
            f'{kern}'
            '<w:sz w:val="21"/>'
            '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
            '<w:name w:val="Normal"/><w:qFormat/></w:style>'
            '</w:styles>')


def make_document(text, jc, page_w_pt, sz_half_pt=21):
    page_w_tw = int(page_w_pt * 20)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>'
            '<w:r>'
            '<w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_half_pt}"/>'
            '</w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t>'
            '</w:r></w:p>'
            '<w:sectPr>'
            f'<w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1700" w:bottom="1440" w:left="1700"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>')


def build_doc(label, with_kern, text, jc, page_w_pt):
    tmp = tempfile.mkdtemp(prefix="m2k_")
    try:
        with zipfile.ZipFile(SRC_COM) as z:
            z.extractall(tmp)
        with open(os.path.join(tmp, "word", "styles.xml"), "w",
                  encoding="utf-8") as f:
            f.write(make_styles(with_kern=with_kern))
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(make_document(text, jc, page_w_pt))
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


# Probe text: 4 commas each between CJK (no Mech 1 trigger).
# Natural width: 27 chars × 10.5pt = 283.5pt
# pgMar 1700+1700 = 3400tw = 170pt. So content = page_w - 170.
# To force overflow: page_w = 280 → content = 110pt. 27 chars don't fit.
# But Mech 2 needs SOME yakumono to compress; can fit only if compression helps.
# Let's use longer text with more yakumono.

PROBE_NO_M1 = "漢漢漢漢、漢漢漢漢、漢漢漢漢、漢漢漢漢、漢漢漢漢"
# B→CJK only, no Mech 1 trigger
PROBE_WITH_M1 = "漢漢漢」（漢漢漢」（漢漢漢」（漢漢漢」（漢漢漢"
# 」（ pairs trigger Mech 1 (B→A)

# Probe sizes: 24-char NO_M1 natural = 252pt; 25-char WITH_M1 natural = 262.5pt.
# Use a content_w that forces small overflow for both.

TESTS = [
    # (label, with_kern, text, jc, pgW)
    # NO Mech 1 trigger (purely 、 between CJK)
    ("M2_noM1_kern_on_slack4",  True,  PROBE_NO_M1, "both",  418),
    ("M2_noM1_kern_off_slack4", False, PROBE_NO_M1, "both",  418),
    # WITH Mech 1 triggers
    ("M2_withM1_kern_on_slack4",  True,  PROBE_WITH_M1, "both",  418),
    ("M2_withM1_kern_off_slack4", False, PROBE_WITH_M1, "both",  418),
    # No overflow ref
    ("M2_noM1_kern_on_no_overflow",   True,  PROBE_NO_M1, "both", 500),
    ("M2_withM1_kern_on_no_overflow", True,  PROBE_WITH_M1, "both", 500),
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for label, with_kern, text, jc, pgW in TESTS:
            path = build_doc(label, with_kern, text, jc, pgW)
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
                        xs.append((t,
                                   float(c.Information(5)),
                                   float(c.Information(6))))
                    except Exception:
                        continue
                d.Close(SaveChanges=False)
                if not xs:
                    advs = []
                    line1 = []
                else:
                    y0 = xs[0][2]
                    line1 = [(c, x) for c, x, y in xs
                             if abs(y - y0) < 0.5]
                    line1_sorted = sorted(line1, key=lambda t: t[1])
                    advs = []
                    for i in range(len(line1_sorted) - 1):
                        advs.append((line1_sorted[i][0],
                                     round(line1_sorted[i + 1][1]
                                           - line1_sorted[i][1], 4)))
                # Look at the 、 advances
                comma_advs = [a for c, a in advs if c == "、"]
                print(f"\n[{label}] kern={with_kern} jc={jc} pgW={pgW}")
                print(f"  n_line1={len(line1)} all advances={advs}")
                print(f"  、 advances: {comma_advs}")
                results[label] = {
                    "with_kern": with_kern, "jc": jc, "page_w_pt": pgW,
                    "n_line1": len(line1),
                    "advances": advs,
                    "comma_advances": comma_advs,
                }
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}")
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
    existing["mech2_kern_isolation_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
