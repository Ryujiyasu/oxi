"""§4.7 trigger — bisect WITHIN styles.xml.

styles.xml is the trigger location. Identify which docDefaults/Normal-style
element is the actual trigger. Use COM doc as base, replace styles.xml with
progressively-stripped versions.
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
OUT_DIR = os.path.abspath("pipeline_data/yakumono_styles_bisect_docs")
RESULT_PATH = os.path.abspath("pipeline_data/ra_manual_measurements.json")


def make_styles(*, with_theme_fonts, with_kern, with_lang,
                 with_ligatures, with_szCs, with_widow,
                 sz_val=21):
    rfonts = ('<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
              if with_theme_fonts else "")
    kern = '<w:kern w:val="2"/>' if with_kern else ""
    sz = f'<w:sz w:val="{sz_val}"/>'
    szCs = f'<w:szCs w:val="24"/>' if with_szCs else ""
    lang = ('<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
            if with_lang else "")
    ligs = ('<w14:ligatures w14:val="standardContextual"/>'
            if with_ligatures else "")
    widow = '<w:widowControl w:val="0"/>' if with_widow else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            f'{rfonts}{kern}{sz}{szCs}{lang}{ligs}'
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
            '<w:name w:val="Normal"/><w:qFormat/>'
            f'<w:pPr>{widow}</w:pPr>'
            '</w:style>'
            '</w:styles>')


VARIANTS = [
    # (label, styles params)
    ("V_full_match_com",   dict(with_theme_fonts=True,  with_kern=True,
                                with_lang=True, with_ligatures=True,
                                with_szCs=True, with_widow=True)),
    ("V_no_theme_fonts",   dict(with_theme_fonts=False, with_kern=True,
                                with_lang=True, with_ligatures=True,
                                with_szCs=True, with_widow=True)),
    ("V_no_kern",          dict(with_theme_fonts=True,  with_kern=False,
                                with_lang=True, with_ligatures=True,
                                with_szCs=True, with_widow=True)),
    ("V_no_lang",          dict(with_theme_fonts=True,  with_kern=True,
                                with_lang=False, with_ligatures=True,
                                with_szCs=True, with_widow=True)),
    ("V_no_ligatures",     dict(with_theme_fonts=True,  with_kern=True,
                                with_lang=True, with_ligatures=False,
                                with_szCs=True, with_widow=True)),
    ("V_no_szCs",          dict(with_theme_fonts=True,  with_kern=True,
                                with_lang=True, with_ligatures=True,
                                with_szCs=False, with_widow=True)),
    ("V_no_widow",         dict(with_theme_fonts=True,  with_kern=True,
                                with_lang=True, with_ligatures=True,
                                with_szCs=True, with_widow=False)),
    # Single-feature variants (each feature alone)
    ("V_only_lang",        dict(with_theme_fonts=False, with_kern=False,
                                with_lang=True,  with_ligatures=False,
                                with_szCs=False, with_widow=False)),
    ("V_only_kern",        dict(with_theme_fonts=False, with_kern=True,
                                with_lang=False, with_ligatures=False,
                                with_szCs=False, with_widow=False)),
    ("V_only_theme",       dict(with_theme_fonts=True,  with_kern=False,
                                with_lang=False, with_ligatures=False,
                                with_szCs=False, with_widow=False)),
    ("V_minimal_sz_only",  dict(with_theme_fonts=False, with_kern=False,
                                with_lang=False, with_ligatures=False,
                                with_szCs=False, with_widow=False)),
]


def build(label, params):
    tmp = tempfile.mkdtemp(prefix="sty_doc_")
    try:
        with zipfile.ZipFile(SRC_COM) as z:
            z.extractall(tmp)
        styles_xml = make_styles(**params)
        with open(os.path.join(tmp, "word", "styles.xml"), "w",
                  encoding="utf-8") as f:
            f.write(styles_xml)
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
        for label, params in VARIANTS:
            path = build(label, params)
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
            results[label] = {"params": params, "advances": advs,
                               "marker": marker}
            print(f"[{label}] {marker} {advs} (params: {params})",
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
    existing["yakumono_styles_bisect_2026-05-02"] = results
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"\nWrote results to {RESULT_PATH}")


if __name__ == "__main__":
    main()
