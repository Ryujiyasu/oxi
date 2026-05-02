"""§4.6.2 round 23 — autoSpaceDE / autoSpaceDN disable + numeral boundary.

Open questions from §4.6.2 (Round 22 prior):
  Q1: Does autoSpaceDN behave identically to autoSpaceDE for digits?
      Probe 漢123 vs 漢abc at fs ∈ {10, 12, 14, 16}.
  Q2: <w:autoSpaceDE w:val="0"/> in settings.xml — disables ALL
      boundary spacing or just CJK→Latin letters?
  Q3: <w:autoSpaceDN w:val="0"/> — disables digit boundary?
  Q4: Combined DE=0 + DN=0 — completely disable?
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\autospace_dn_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\autospace_dn.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"
LATIN_FONT = "Times New Roman"


def make_doc_xml(probe, sz_val, page_w_tw, margin_tw):
    """Single run, mixed CJK + Latin."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="left"/>'
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


def make_styles_xml(sz_val):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml(de_val=None, dn_val=None):
    """de_val/dn_val: None = no override (Word default ON), 0 = explicitly OFF."""
    inner = ""
    if de_val is not None:
        inner += f'<w:autoSpaceDE w:val="{de_val}"/>'
    if dn_val is not None:
        inner += f'<w:autoSpaceDN w:val="{dn_val}"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'{inner}'
            '</w:settings>')


def make_docx(label, probe, sz_val, de_val=None, dn_val=None):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((400 + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(probe, sz_val, page_w_tw, margin_tw)
    styles_xml = make_styles_xml(sz_val)
    settings_xml = make_settings_xml(de_val, dn_val)
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
                    xs.append((t, float(c.Information(5)), float(c.Information(6)),
                               float(c.Font.Size if c.Font.Size else 0)))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        char_advances = []
        for i in range(len(xs) - 1):
            t = xs[i][0]
            adv = round(xs[i+1][1] - xs[i][1], 3)
            char_advances.append({"ch": t, "adv": adv, "sz": xs[i][3]})
        return {
            "n_chars": len(xs),
            "char_advances": char_advances,
        }
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}

    # Probes:
    #   "漢a" — kana→Latin letter (DE flag controls)
    #   "漢1" — kana→Latin digit (DN flag controls)
    #   "漢字abc1" — multi-boundary (kana, kana, letter, letter, letter, digit)
    #   "a漢" — Latin letter→kana
    #   "1漢" — Latin digit→kana
    probes = [
        ("Q1_kana_letter", "漢a", "kana→letter"),
        ("Q1_kana_digit",  "漢1", "kana→digit"),
        ("Q1_letter_kana", "a漢", "letter→kana"),
        ("Q1_digit_kana",  "1漢", "digit→kana"),
        ("Q2_mixed",       "漢a1漢", "multi-boundary"),
    ]

    # Settings combinations: (de, dn) — None = default (ON)
    settings = [
        ("def",   None, None),  # both default ON
        ("DE0",     0,  None),  # DE=0
        ("DN0",   None,    0),  # DN=0
        ("D00",     0,    0),   # both 0
    ]

    fs_pt = 12.0
    sz_val = 24

    for probe_label, probe, desc in probes:
        print(f"\n=== {probe_label}: {probe!r} ({desc}) at fs={fs_pt}pt ===")
        out[probe_label] = {"probe": probe, "desc": desc, "fs": fs_pt, "tests": []}
        for s_label, de_val, dn_val in settings:
            label = f"{probe_label}_{s_label}"
            try:
                p = make_docx(label, probe, sz_val, de_val, dn_val)
            except Exception as e:
                out[probe_label]["tests"].append({"settings": s_label, "build_error": str(e)})
                continue
            kill_word()
            try:
                r = measure_one(p)
            except Exception as e:
                r = {"measure_error": str(e)}
                kill_word()
            entry = {"settings": s_label, "de_val": de_val, "dn_val": dn_val, **r}
            out[probe_label]["tests"].append(entry)
            advs = entry.get("char_advances", [])
            adv_str = ' '.join(f"{c['ch']}={c['adv']:.1f}" for c in advs[:6])
            print(f"  settings={s_label:>4} (DE={de_val} DN={dn_val}): {adv_str}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== SUMMARY ==========")
    print("Per probe, advance of FIRST char (boundary char):")
    print(f"{'probe':>20} {'def':>8} {'DE0':>8} {'DN0':>8} {'D00':>8}")
    for probe_label, probe, _ in probes:
        info = out.get(probe_label, {})
        row = {}
        for t in info.get("tests", []):
            advs = t.get("char_advances", [])
            if advs:
                row[t["settings"]] = advs[0]["adv"]
        print(f"{probe_label:>20} "
              f"{row.get('def','?'):>8} {row.get('DE0','?'):>8} "
              f"{row.get('DN0','?'):>8} {row.get('D00','?'):>8}")
    print("\n  natural CJK (12pt) = 12.0; natural digit = ~6.0; natural letter = varies")
    print("  Round 22 (§4.6.2) at fs=12: extra = +3.0pt → boundary char adv = 12+3 or 6+3 etc")


if __name__ == "__main__":
    main()
