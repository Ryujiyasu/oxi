"""§4.7 — Test if w:kern affects other mechanisms beyond Mech 1.

We've pinpointed w:kern as the gate for Mech 1 (Type A/B/C adjacency
compression). Now test:
  T1: Does w:kern off disable Mech 2 (justify-time slack distribution)?
  T2: Does w:kern on activate §4.6.3 (CJK-adjacent space widening)?
  T3: Does w:kern affect §4.6.2 (kana→Latin alphanumeric autoSpaceDE)?

Method: clone COM-generated docx (which has kern). Synthesize 4 styles.xml
variants with kern present/absent. For each, test Mech 2 (overflow line),
§4.6.3 (Latin space before CJK), §4.6.2 (kana then Latin alpha).
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
OUT_DIR = os.path.abspath("pipeline_data/kern_other_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")


def make_styles(with_kern=True):
    kern = '<w:kern w:val="2"/>' if with_kern else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            f'<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
            f'{kern}'
            '<w:sz w:val="21"/>'
            '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
            '<w:name w:val="Normal"/><w:qFormat/></w:style>'
            '</w:styles>')


def make_document(text, jc, page_w_pt, font_size_half_pt=21):
    page_w_tw = int(page_w_pt * 20)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>'
            '<w:r>'
            '<w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
            f'<w:sz w:val="{font_size_half_pt}"/>'
            '</w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t>'
            '</w:r></w:p>'
            '<w:sectPr>'
            f'<w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1700" w:bottom="1440" w:left="1700"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>')


def build_doc(label, with_kern, text, jc, page_w_pt, sz_half_pt=21):
    tmp = tempfile.mkdtemp(prefix="kern_other_")
    try:
        with zipfile.ZipFile(SRC_COM) as z:
            z.extractall(tmp)
        # Replace styles.xml
        with open(os.path.join(tmp, "word", "styles.xml"), "w",
                  encoding="utf-8") as f:
            f.write(make_styles(with_kern=with_kern))
        # Replace document.xml
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(make_document(text, jc, page_w_pt, sz_half_pt))
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


# Tests:
# T1 Mech 2: jc=both with overflow. Use 」（ pairs that compress under
# Mech 1 (when kern on) AND additional candidates for Mech 2.
# Adjacent 」（ pair = B→A trigger.
T1_TEXT = "漢漢漢」（漢漢漢」（漢漢漢」（漢漢漢"
T1_JC = "both"
T1_PAGE_W_PT = 380  # content_w = 210pt — moderate overflow
# Natural width: 19 × 10.5 = 199.5pt. With 3 」 compressed (Mech 1) at
# 5.5pt each, saves 15pt → 184.5pt fits content_w 210pt with slack 25.5pt.
# Mech 2 may distribute that slack to （s.

# T2 §4.6.3: Latin space before CJK
T2_TEXT = "Foo は M"
T2_JC = "left"
T2_PAGE_W_PT = 595.3

# T3 §4.6.2: kana DIRECTLY followed by Latin alphanumeric
T3_TEXT = "はMでs"
T3_JC = "left"
T3_PAGE_W_PT = 595.3

TESTS = [
    ("T1_mech2_kern_on",  True,  T1_TEXT, T1_JC, T1_PAGE_W_PT),
    ("T1_mech2_kern_off", False, T1_TEXT, T1_JC, T1_PAGE_W_PT),
    ("T2_4_6_3_kern_on",  True,  T2_TEXT, T2_JC, T2_PAGE_W_PT),
    ("T2_4_6_3_kern_off", False, T2_TEXT, T2_JC, T2_PAGE_W_PT),
    ("T3_4_6_2_kern_on",  True,  T3_TEXT, T3_JC, T3_PAGE_W_PT),
    ("T3_4_6_2_kern_off", False, T3_TEXT, T3_JC, T3_PAGE_W_PT),
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
                results[label] = {
                    "with_kern": with_kern, "text": text, "jc": jc,
                    "page_w_pt": pgW,
                    "n_line1": len(line1) if xs else 0,
                    "advances": advs,
                }
                print(f"\n[{label}] kern={with_kern} text={text}")
                print(f"  n_line1={len(line1) if xs else 0} advances={advs}")
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
    existing["kern_other_mechanisms_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
