"""§4.7c Mech 3 — confirm characterSpacingControl value sensitivity.

Round 1 found cSC=compressPunctuation is the sole Mech 3 trigger.
Round 2 (this script): test all 3 ECMA-376 valid values:
  - doNotCompress (default)
  - compressPunctuation
  - compressPunctuationAndJapaneseKana

Plus combination test: cSC=compressPunctuation + kern removed → both
Mech 1 and Mech 3 status verified.
"""
import json, os, sys, re, time, zipfile, shutil, tempfile
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC_REAL = os.path.abspath(
    r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\7f272a2dfd3b_index-21.docx")
OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\mech3_csc_docs")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\mech3_csc_values.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("（「『【〔｛〈《［）」』】〕｝〉》］、。，．—")

PROBE_TEXT = (
    "卸売市場法第６条第１項（第14条において準用する同法第６条第１項）"
    "の規定により、中央卸売市場（地方卸売市場）に係る認定事項の変更について"
    "認定を受けたいので、次のとおり関係書類を添えて申請します。"
)


def make_doc_xml(text, jc):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
            '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"'
            ' w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def replace_csc(s, value):
    """Replace cSC value (or remove element if value is None)."""
    if value is None:
        return re.sub(r'<w:characterSpacingControl[^/]*/>', '', s)
    return re.sub(
        r'<w:characterSpacingControl[^/]*?w:val="[^"]*"/>',
        f'<w:characterSpacingControl w:val="{value}"/>',
        s
    )


def remove_kern(s):
    return re.sub(
        r'(<w:docDefaults>.*?<w:rPrDefault>.*?<w:rPr>)(.*?)(</w:rPr>)',
        lambda m: m.group(1) + re.sub(r'<w:kern[^/]*/>', '', m.group(2)) + m.group(3),
        s, flags=re.S, count=1
    )


def make_variant(label, csc_value=None, kern_removed=False, jc="left"):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    tmp = tempfile.mkdtemp(prefix="csc_")
    try:
        with zipfile.ZipFile(SRC_REAL) as z:
            z.extractall(tmp)
        with open(os.path.join(tmp, "word", "document.xml"), "w", encoding="utf-8") as f:
            f.write(make_doc_xml(PROBE_TEXT, jc))
        # settings.xml
        sp = os.path.join(tmp, "word", "settings.xml")
        with open(sp, "r", encoding="utf-8") as f:
            s = f.read()
        if csc_value is not None or csc_value == "":  # explicit
            s = replace_csc(s, csc_value if csc_value else None)
        with open(sp, "w", encoding="utf-8") as f:
            f.write(s)
        # styles.xml
        if kern_removed:
            stp = os.path.join(tmp, "word", "styles.xml")
            with open(stp, "r", encoding="utf-8") as f:
                s2 = f.read()
            s2 = remove_kern(s2)
            with open(stp, "w", encoding="utf-8") as f:
                f.write(s2)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for fn in files:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
        return out_path
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def measure(word, path):
    d = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.2)
    try:
        chars = d.Range().Characters
        xs = []
        for ci in range(1, chars.Count + 1):
            try:
                c = chars(ci)
                t = c.Text
                if t in ("\r", "\x07"):
                    continue
                xs.append((t, float(c.Information(5)), float(c.Information(6)),
                           float(c.Font.Size if c.Font.Size else 0)))
            except Exception:
                continue
    finally:
        try: d.Close(SaveChanges=False)
        except: pass
    if not xs: return {"error": "no chars"}
    lines_b = {}
    for t, x, y, sz in xs:
        ykey = round(y, 0)
        lines_b.setdefault(ykey, []).append((t, x, y, sz))
    n_yak_comp = 0
    n_yak_half = 0  # half-width = Mech 1
    n_yak_partial = 0  # partial = Mech 2/3
    detail_first = []
    for ykey in sorted(lines_b.keys()):
        items = sorted(lines_b[ykey], key=lambda v: v[1])
        for i in range(len(items) - 1):
            t, a, s = items[i][0], round(items[i+1][1] - items[i][1], 3), items[i][3]
            if t in YAKUMONO and s > 0 and a < s * 0.99:
                n_yak_comp += 1
                if a < s * 0.6:
                    n_yak_half += 1   # half-width = Mech 1
                else:
                    n_yak_partial += 1  # partial = Mech 2/3
                if len(detail_first) < 8:
                    detail_first.append((t, round(a, 2), round(s, 1),
                                          "M1half" if a < s * 0.6 else "M2/3partial"))
    return {
        "n_yak_compressed_total": n_yak_comp,
        "n_mech1_half": n_yak_half,
        "n_mech23_partial": n_yak_partial,
        "fires_mech1": n_yak_half > 0,
        "fires_mech23": n_yak_partial > 0,
        "detail": detail_first,
    }


VARIANTS = [
    # (label, csc_value, kern_removed, jc)
    ("CSC0_baseline_compress_left",      "compressPunctuation",          False, "left"),
    ("CSC1_doNotCompress_left",          "doNotCompress",                False, "left"),
    ("CSC2_compressKana_left",           "compressPunctuationAndJapaneseKana", False, "left"),
    ("CSC3_default_no_element_left",     "",  False, "left"),  # remove element entirely

    # Cross with kern
    ("KX0_compress_kernYes_left",        "compressPunctuation",          False, "left"),
    ("KX1_compress_kernNo_left",         "compressPunctuation",          True,  "left"),
    ("KX2_doNotCompress_kernYes_left",   "doNotCompress",                False, "left"),
    ("KX3_doNotCompress_kernNo_left",    "doNotCompress",                True,  "left"),

    # jc=both crosscheck
    ("JC0_compress_both",                "compressPunctuation",          False, "both"),
    ("JC1_doNotCompress_both",           "doNotCompress",                False, "both"),
]


def main():
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    out = {}
    try:
        for label, csc, kern_rm, jc in VARIANTS:
            try:
                p = make_variant(label, csc, kern_rm, jc)
            except Exception as e:
                out[label] = {"build_error": str(e)}
                print(f"[{label}] BUILD ERR: {e}")
                continue
            try:
                r = measure(word, p)
                out[label] = {"csc": csc, "kern_removed": kern_rm, "jc": jc, **r}
                m1 = "M1+" if r.get("fires_mech1") else "M1-"
                m23 = "M23+" if r.get("fires_mech23") else "M23-"
                print(f"[{label:<40}] csc={csc!r:35s} kern={'no' if kern_rm else 'yes'} jc={jc:5s} {m1}/{m23}  comp={r.get('n_yak_compressed_total')} (half={r.get('n_mech1_half')} partial={r.get('n_mech23_partial')})")
            except Exception as e:
                out[label] = {"measure_error": str(e)}
                print(f"[{label}] MEASURE ERR: {e}")
    finally:
        try: word.Quit()
        except: pass

    os.makedirs(os.path.dirname(RESULT), exist_ok=True)
    with open(RESULT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT}")


if __name__ == "__main__":
    main()
