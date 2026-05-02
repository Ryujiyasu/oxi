"""§4.7b Mech 2 vs §4.7c Mech 3 — verify whether they are the same mechanism.

§4.7b claimed Mech 2 fires only at jc=both, characterized via synthesized
minimal OOXML that lacked settings.xml entirely. ECMA-376 default for
characterSpacingControl is "doNotCompress", so by §4.7c findings those
synthesized docs SHOULD NOT have fired Mech 3.

But §4.7b reports they DID fire. Either:
  (A) Mech 2 and Mech 3 are TWO independent mechanisms with different
      triggers (Mech 2 = overflow+jc=both, Mech 3 = cSC=compressPunct).
  (B) §4.7b synthesized docs accidentally inherited cSC from Word
      defaults somehow, contradicting the "minimal" claim.
  (C) Word treats absence of settings.xml differently from
      cSC=doNotCompress.

This script reproduces §4.7b's Mech 2 setup with explicit cSC values
to disambiguate.

Probe template: 漢×5 + 「 + 漢×13 (single yak mid-line)
Variant: page width that forces overflow (slack = 4pt at MS Mincho 12pt).
"""
import json, os, sys, time, zipfile
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\m2_csc_verify_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\m2_csc_verify.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("（「『【〔｛〈《［）」』】〕｝〉》］、。，．—")

# 20-char probe with 1 yak at pos 6, content_w=236pt → slack=4pt at 12pt
PROBE = "漢漢漢漢漢「漢漢漢漢漢漢漢漢漢漢漢漢漢漢"   # 20 chars total


def make_doc_xml_inline_settings(jc, csc_value=None, font_size_x2=24):
    """Build a complete docx with explicit settings.xml content
    controlling cSC value."""
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body><w:p>'
        f'<w:pPr><w:jc w:val="{jc}"/>'
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        '</w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
        ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/>'
        f'<w:sz w:val="{font_size_x2}"/></w:rPr>'
        f'<w:t>{PROBE}</w:t></w:r></w:p>'
        '<w:sectPr>'
        f'<w:pgSz w:w="{int(236*20+170*20)}" w:h="16838"/>'
        f'<w:pgMar w:top="1134" w:right="{170*10}" w:bottom="1134" w:left="{170*10}"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr></w:body></w:document>'
    )
    if csc_value:
        settings_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:characterSpacingControl w:val="{csc_value}"/>'
            '</w:settings>'
        )
    elif csc_value == "":
        # Empty settings.xml (no cSC element)
        settings_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        )
    else:
        settings_xml = None
    return doc_xml, settings_xml


def make_docx(label, jc, csc_value):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    doc_xml, settings_xml = make_doc_xml_inline_settings(jc, csc_value)
    has_settings = settings_xml is not None
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        + (
            '<Override PartName="/word/settings.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
            if has_settings else ''
        )
        + '</Types>'
    )
    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>'
    )
    if has_settings:
        doc_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
            ' Target="settings.xml"/>'
            '</Relationships>'
        )
    else:
        doc_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
        if has_settings:
            z.writestr("word/settings.xml", settings_xml)
    return out_path


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
    n_yak_total = 0
    yak_detail = []
    for ykey in sorted(lines_b.keys()):
        items = sorted(lines_b[ykey], key=lambda v: v[1])
        for i in range(len(items) - 1):
            t = items[i][0]
            a = round(items[i+1][1] - items[i][1], 3)
            sz = items[i][3]
            if t in YAKUMONO:
                n_yak_total += 1
                yak_detail.append((t, a, sz))
                if sz > 0 and a < sz * 0.99:
                    n_yak_comp += 1
    return {
        "n_yak_total": n_yak_total,
        "n_yak_compressed": n_yak_comp,
        "fires": n_yak_comp > 0,
        "yak_detail": [(t, round(a, 2), round(s, 1)) for t, a, s in yak_detail],
    }


# Test matrix: 4 csc states × 2 jc
VARIANTS = []
for csc in ["compressPunctuation", "doNotCompress", None, "EMPTY"]:
    csc_short = {"compressPunctuation": "comp", "doNotCompress": "noComp",
                 None: "noSettings", "EMPTY": "emptySettings"}[csc]
    csc_xml_value = "" if csc == "EMPTY" else csc
    for jc in ["both", "left"]:
        VARIANTS.append((f"V_{csc_short}_{jc}", csc_xml_value, jc))


def kill_word():
    import subprocess
    subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
    time.sleep(2)


def main():
    out = {}
    for label, csc_value, jc in VARIANTS:
        try:
            p = make_docx(label, jc, csc_value)
        except Exception as e:
            out[label] = {"build_error": str(e)}
            print(f"[{label}] BUILD ERR: {e}")
            continue
        kill_word()
        try:
            word = w32.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
        except Exception as e:
            out[label] = {"start_error": str(e)}
            continue
        try:
            r = measure(word, p)
            out[label] = {"csc_value": csc_value, "jc": jc, **r}
            fire = "FIRE" if r.get("fires") else "no  "
            print(f"[{label:<28}] csc={csc_value!r:30s} jc={jc:<5}  {fire}  yak_comp={r.get('n_yak_compressed','?')}/{r.get('n_yak_total','?')}")
        except Exception as e:
            out[label] = {"measure_error": str(e)}
            print(f"[{label}] MEASURE ERR: {e}")
        finally:
            try: word.Quit()
            except: pass

    os.makedirs(os.path.dirname(RESULT), exist_ok=True)
    with open(RESULT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\n=== Verdict ===")
    print(f"{'state':<20} | {'jc=both':<10} | {'jc=left':<10}")
    print("-" * 50)
    for csc_short in ["comp", "noComp", "noSettings", "emptySettings"]:
        b = "FIRE" if out.get(f"V_{csc_short}_both", {}).get("fires") else "no  "
        l = "FIRE" if out.get(f"V_{csc_short}_left", {}).get("fires") else "no  "
        print(f"{csc_short:<20} | {b:<10} | {l:<10}")


if __name__ == "__main__":
    main()
