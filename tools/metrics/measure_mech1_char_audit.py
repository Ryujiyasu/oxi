"""§4.7 round 12 — Mech 1 trigger audit for individual Type A/B chars.

§4.7 lists 11 Type A + 13 Type B chars. Spec says:
  Type A: compresses when preceded by Type A (A→A pair)
  Type B: compresses when followed by Type A or B (B→A, B→B)

Verify each individual char compresses under expected trigger pair.

Trigger conditions for Mech 1:
  - <w:kern> in docDefaults rPrDefault rPr (round 6 finding)
  - char in correct position with correct neighbor

Probe: 8-char line, no overflow needed (Mech 1 fires regardless).

Test set (per spec §4.7 lines 614-619, excluding smart quotes done in
round 11 and em-dash whose membership is glyph-metric):

Type A: （ U+FF08, 「 U+300C, 『 U+300E, 【 U+3010, 〔 U+3014,
        ｛ U+FF5B, 〈 U+3008, 《 U+300A, ［ U+FF3B
Type B: ） U+FF09, 」 U+300D, 』 U+300F, 】 U+3011, 〕 U+3015,
        ｝ U+FF5D, 〉 U+3009, 》 U+300B, ］ U+FF3D,
        、 U+3001, 。 U+3002, ， U+FF0C, ． U+FF0E
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\mech1_audit_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\mech1_char_audit.json")
os.makedirs(OUT_DIR, exist_ok=True)

TYPE_A_CHARS = [
    ("LParen",     "（", "U+FF08"),
    ("LCornerB",   "「", "U+300C"),
    ("LDblCornerB","『", "U+300E"),
    ("LBlackB",    "【", "U+3010"),
    ("LTortoise",  "〔", "U+3014"),
    ("LBrace",     "｛", "U+FF5B"),
    ("LAngleB",    "〈", "U+3008"),
    ("LDblAngleB", "《", "U+300A"),
    ("LSqB",       "［", "U+FF3B"),
]

TYPE_B_CHARS = [
    ("RParen",     "）", "U+FF09"),
    ("RCornerB",   "」", "U+300D"),
    ("RDblCornerB","』", "U+300F"),
    ("RBlackB",    "】", "U+3011"),
    ("RTortoise",  "〕", "U+3015"),
    ("RBrace",     "｝", "U+FF5D"),
    ("RAngleB",    "〉", "U+3009"),
    ("RDblAngleB", "》", "U+300B"),
    ("RSqB",       "］", "U+FF3D"),
    ("Comma",      "、", "U+3001"),
    ("Period",     "。", "U+3002"),
    ("FullComma",  "，", "U+FF0C"),
    ("FullPeriod", "．", "U+FF0E"),
]


def make_doc_xml(probe, jc, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/>'
            '<w:sz w:val="24"/></w:rPr>'
            f'<w:t>{probe}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml():
    """Include kern in docDefaults for Mech 1 trigger."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault>'
            '<w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/>'
            '<w:kern w:val="2"/>'
            '<w:sz w:val="24"/>'
            '</w:rPr>'
            '</w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults>'
            '</w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '</w:settings>')


def make_docx(label, probe):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int(280 * 20)   # 280pt page → ~110pt content
    margin_tw = int(85 * 20)
    doc_xml = make_doc_xml(probe, "both", page_w_tw, margin_tw)
    styles_xml = make_styles_xml()
    settings_xml = make_settings_xml()
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '</Types>'
    )
    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>'
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
        ' Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
        ' Target="settings.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", styles_xml)
        z.writestr("word/settings.xml", settings_xml)
    return out_path


def kill_word():
    subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
    time.sleep(2)


def measure_one(path):
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        d = word.Documents.Open(str(path), ReadOnly=True)
        time.sleep(0.2)
        try:
            chars = d.Range().Characters
            xs = []
            for ci in range(1, chars.Count + 1):
                try:
                    c = chars(ci)
                    t = c.Text
                    if t in ("\r", "\x07"): continue
                    xs.append((t, float(c.Information(5)), float(c.Information(6)),
                               float(c.Font.Size if c.Font.Size else 0)))
                except Exception: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        # Single-line, sort by x
        xs.sort(key=lambda v: v[1])
        advs = []
        for i in range(len(xs) - 1):
            t = xs[i][0]
            adv = round(xs[i+1][1] - xs[i][1], 3)
            sz = xs[i][3]
            advs.append({"ch": t, "adv": adv, "sz": sz,
                         "ratio": round(adv/sz, 4) if sz > 0 else None,
                         "compressed": adv < sz * 0.95 if sz > 0 else False})
        return {"n_chars": len(xs), "advances": advs}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}

    # Suite A: Type A → Type A (A_i preceded by （)
    print("\n========== Suite A: A→A (each Type A preceded by （) ==========")
    for label, ch, uni in TYPE_A_CHARS:
        probe = "漢漢漢（" + ch + "漢漢漢"   # （ at pos 4, char ch at pos 5
        try:
            p = make_docx(f"A_AA_{label}", probe)
        except Exception as e:
            out[f"A_AA_{label}"] = {"build_error": str(e)}
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        # Find ch advance
        target = None
        if "advances" in r:
            for adv_info in r["advances"]:
                if adv_info["ch"] == ch:
                    target = adv_info
                    break
        out[f"A_AA_{label}"] = {"char": label, "unicode": uni, "probe": probe,
                                 "char_test": ch, "target": target}
        if target:
            comp = "FIRES" if target["compressed"] else "no   "
            print(f"  ({label:<14}{uni}) adv={target['adv']:.2f} ratio={target['ratio']} {comp}")
        else:
            print(f"  ({label:<14}{uni}) ERROR or not found")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    # Suite B: Type B → Type B (each Type B followed by ）)
    print("\n========== Suite B: B→B (each Type B followed by ）) ==========")
    for label, ch, uni in TYPE_B_CHARS:
        probe = "漢漢漢" + ch + "）漢漢漢"   # char ch at pos 4, ） at pos 5
        try:
            p = make_docx(f"B_BB_{label}", probe)
        except Exception as e:
            out[f"B_BB_{label}"] = {"build_error": str(e)}
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        target = None
        if "advances" in r:
            for adv_info in r["advances"]:
                if adv_info["ch"] == ch:
                    target = adv_info
                    break
        out[f"B_BB_{label}"] = {"char": label, "unicode": uni, "probe": probe,
                                 "char_test": ch, "target": target}
        if target:
            comp = "FIRES" if target["compressed"] else "no   "
            print(f"  ({label:<14}{uni}) adv={target['adv']:.2f} ratio={target['ratio']} {comp}")
        else:
            print(f"  ({label:<14}{uni}) ERROR")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    # Suite C: Control — each Type B between CJK (no Mech 1 expected)
    print("\n========== Suite C (control): each Type B between CJK ==========")
    for label, ch, uni in TYPE_B_CHARS[:5]:   # test 5 representative
        probe = "漢漢漢" + ch + "漢漢漢漢"
        try:
            p = make_docx(f"C_CTL_{label}", probe)
        except Exception as e:
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            kill_word()
            continue
        target = None
        if "advances" in r:
            for adv_info in r["advances"]:
                if adv_info["ch"] == ch:
                    target = adv_info
                    break
        out[f"C_CTL_{label}"] = {"char": label, "unicode": uni, "probe": probe,
                                  "char_test": ch, "target": target}
        if target:
            comp = "FIRES" if target["compressed"] else "no"
            print(f"  ({label:<14}{uni}) adv={target['adv']:.2f} {comp}  (expected: no)")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    # Summary
    print("\n========== SUMMARY ==========")
    n_a_pass = sum(1 for k, v in out.items() if k.startswith("A_AA_") and v.get("target", {}).get("compressed"))
    n_a_total = sum(1 for k in out if k.startswith("A_AA_"))
    n_b_pass = sum(1 for k, v in out.items() if k.startswith("B_BB_") and v.get("target", {}).get("compressed"))
    n_b_total = sum(1 for k in out if k.startswith("B_BB_"))
    n_c_pass = sum(1 for k, v in out.items() if k.startswith("C_CTL_") and not v.get("target", {}).get("compressed"))
    n_c_total = sum(1 for k in out if k.startswith("C_CTL_"))
    print(f"Suite A (A→A): {n_a_pass}/{n_a_total} chars FIRE Mech 1")
    print(f"Suite B (B→B): {n_b_pass}/{n_b_total} chars FIRE Mech 1")
    print(f"Suite C (control): {n_c_pass}/{n_c_total} chars correctly DON'T compress")


if __name__ == "__main__":
    main()
