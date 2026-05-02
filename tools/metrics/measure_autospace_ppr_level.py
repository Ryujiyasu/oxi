"""§4.6.2 round 24 — autoSpaceDE/DN at paragraph level (pPr).

Round 23 found settings.xml-level flags had no effect. ECMA-376 says
autoSpaceDE/DN are paragraph properties (CT_PPrBase). Test pPr-level
to see if Word actually disables boundary there.

Probes: 漢a (kana→letter), 漢1 (kana→digit), 1漢 (digit→kana)
Settings (per paragraph):
  def  — no override
  DE0  — pPr/autoSpaceDE val=0
  DN0  — pPr/autoSpaceDN val=0
  D00  — both
  STY  — same flags via styles docDefaults pPrDefault
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\autospace_ppr_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\autospace_ppr.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"
LATIN_FONT = "Times New Roman"


def make_doc_xml(probe, sz_val, page_w_tw, margin_tw, ppr_flags=""):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="left"/>{ppr_flags}'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t xml:space="preserve">{probe}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml(sz_val, sty_flags=""):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            f'<w:pPrDefault><w:pPr>{sty_flags}</w:pPr></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '</w:settings>')


def make_docx(label, probe, sz_val, ppr_flags="", sty_flags=""):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((400 + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(probe, sz_val, page_w_tw, margin_tw, ppr_flags)
    styles_xml = make_styles_xml(sz_val, sty_flags)
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
                    if t in ("\r","\x07"): continue
                    xs.append((t, float(c.Information(5)), float(c.Information(6))))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        char_advances = []
        for i in range(len(xs) - 1):
            t = xs[i][0]
            adv = round(xs[i+1][1] - xs[i][1], 3)
            char_advances.append({"ch": t, "adv": adv})
        return {"n_chars": len(xs), "char_advances": char_advances}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    sz_val = 24
    fs_pt = 12.0

    probes = [
        ("kana_letter", "漢a"),
        ("kana_digit",  "漢1"),
        ("digit_kana",  "1漢"),
        ("letter_kana", "a漢"),
    ]

    # Settings: (label, ppr_flags, sty_flags)
    settings = [
        ("def", "", ""),
        ("ppr_DE0",  '<w:autoSpaceDE w:val="0"/>', ""),
        ("ppr_DN0",  '<w:autoSpaceDN w:val="0"/>', ""),
        ("ppr_D00",  '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>', ""),
        ("sty_D00",  "", '<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>'),
    ]

    for probe_label, probe in probes:
        print(f"\n=== {probe_label}: {probe!r} ===")
        out[probe_label] = {"probe": probe, "tests": []}
        for s_label, ppr_flags, sty_flags in settings:
            label = f"{probe_label}_{s_label}"
            try:
                p = make_docx(label, probe, sz_val, ppr_flags, sty_flags)
            except Exception as e:
                out[probe_label]["tests"].append({"settings": s_label, "build_error": str(e)})
                continue
            kill_word()
            try:
                r = measure_one(p)
            except Exception as e:
                r = {"measure_error": str(e)}
                kill_word()
            entry = {"settings": s_label, **r}
            out[probe_label]["tests"].append(entry)
            advs = entry.get("char_advances", [])
            adv_str = ' '.join(f"{c['ch']}={c['adv']:.1f}" for c in advs[:4])
            print(f"  {s_label:>9}: {adv_str}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== SUMMARY ==========")
    print(f"{'probe':>15} {'def':>10} {'pprDE0':>10} {'pprDN0':>10} {'pprD00':>10} {'styD00':>10}")
    for probe_label, probe in probes:
        info = out.get(probe_label, {})
        row = {}
        for t in info.get("tests", []):
            advs = t.get("char_advances", [])
            if advs:
                row[t["settings"]] = advs[0]["adv"]
        print(f"{probe_label:>15} "
              f"{row.get('def','?'):>10} {row.get('ppr_DE0','?'):>10} "
              f"{row.get('ppr_DN0','?'):>10} {row.get('ppr_D00','?'):>10} "
              f"{row.get('sty_D00','?'):>10}")


if __name__ == "__main__":
    main()
