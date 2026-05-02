"""§4.7 — Test w:kern semantics:
1. Does `<w:kern w:val="0"/>` enable or disable?
2. Do val=1, val=100, val=2000 give same behavior?
3. Can pPr override docDefaults kern?
4. Can run rPr override docDefaults kern?

Probe: 漢」（漢 (B→A pair, 」 should compress under Mech 1 if kern enabled).
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
OUT_DIR = os.path.abspath("pipeline_data/kern_value_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")

PROBE = "漢」（漢"


def make_styles(doc_default_kern_val):
    """doc_default_kern_val: None to omit, int to set w:val=N"""
    if doc_default_kern_val is None:
        kern = ""
    else:
        kern = f'<w:kern w:val="{doc_default_kern_val}"/>'
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


def make_document(text, ppr_kern_val=None, run_rpr_kern_val=None):
    ppr_extra = ""
    if ppr_kern_val is not None:
        ppr_extra = ('<w:rPr>'
                      f'<w:kern w:val="{ppr_kern_val}"/>'
                      '</w:rPr>')
    run_kern = ""
    if run_rpr_kern_val is not None:
        run_kern = f'<w:kern w:val="{run_rpr_kern_val}"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="left"/>{ppr_extra}</w:pPr>'
            '<w:r>'
            '<w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
            f'{run_kern}'
            '<w:sz w:val="21"/>'
            '</w:rPr>'
            f'<w:t>{text}</w:t>'
            '</w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1700" w:bottom="1440" w:left="1700"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>')


def build_doc(label, dd_kern, ppr_kern, run_kern):
    tmp = tempfile.mkdtemp(prefix="kv_")
    try:
        with zipfile.ZipFile(SRC_COM) as z:
            z.extractall(tmp)
        with open(os.path.join(tmp, "word", "styles.xml"), "w",
                  encoding="utf-8") as f:
            f.write(make_styles(dd_kern))
        with open(os.path.join(tmp, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(make_document(PROBE, ppr_kern, run_kern))
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


# (label, dd_kern, ppr_kern, run_kern)
TESTS = [
    # Q1: Does kern val=0 enable or disable?
    ("Q1_dd_no_kern",        None,  None,  None),  # control: no kern element
    ("Q1_dd_kern_val_0",     0,     None,  None),  # val=0
    ("Q1_dd_kern_val_1",     1,     None,  None),
    ("Q1_dd_kern_val_2",     2,     None,  None),
    ("Q1_dd_kern_val_100",   100,   None,  None),
    ("Q1_dd_kern_val_1000",  1000,  None,  None),
    ("Q1_dd_kern_val_2000",  2000,  None,  None),
    ("Q1_dd_kern_val_99999", 99999, None,  None),
    # Q2: pPr override
    ("Q2_dd_no_ppr_kern2",  None, 2,    None),  # turn on via pPr
    ("Q2_dd_kern2_ppr_no",  2,    0,    None),  # turn off via pPr (val=0)
    # Q3: run rPr override
    ("Q3_dd_no_run_kern2",  None, None, 2),     # turn on via run
    ("Q3_dd_kern2_run_no",  2,    None, 0),     # turn off via run
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = {}
    try:
        for label, dd_kern, ppr_kern, run_kern in TESTS:
            path = build_doc(label, dd_kern, ppr_kern, run_kern)
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
            kakko_adv = (advs[1][1] if (isinstance(advs, list)
                                         and len(advs) >= 2) else None)
            marker = ("?" if kakko_adv is None
                       else ("COMPRESSED" if kakko_adv < 8 else "FULL"))
            results[label] = {
                "dd_kern": dd_kern, "ppr_kern": ppr_kern,
                "run_kern": run_kern,
                "advances": advs, "marker": marker,
            }
            print(f"[{label}] dd={dd_kern} ppr={ppr_kern} run={run_kern} "
                  f"→ {marker} {advs}", flush=True)
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
    existing["kern_value_override_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
