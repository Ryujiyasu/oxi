"""§4.7b round 19 — pure-yak line cap behavior.

Round 16 confirmed cap = last CJK run's fs/2. Question: what about
lines with NO CJK? (e.g., all-yak emoji-style decorative text)

Probe variants:
  A: 「」「」「」「」「」「」「」「」「」「」「」「」 (24 chars: 12 A + 12 B alternating)
  B: 「漢」漢「漢」漢「漢」漢「漢」漢「漢」漢「漢」漢「漢」漢 (24 chars: 12 yak + 12 CJK alternating)

Suite A reveals: pure-yak cap?
Suite B (control): yak heavy line with CJK, expected cap=fs/2

After Mech 1 fires on alternating 「」:
  - Each 」 followed by 「 → 」 compresses (B→A trigger)
  - Each 「 preceded by 」 → 「 doesn't compress (A only by A)
  - Result: 」 = 6pt, 「 = 12pt (for 12pt fs)
  - Pure-yak line natural after Mech 1 = 6*12+6*6 = 108pt for 12 chars
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\pure_yak_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\pure_yak.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「」")
FONT = "ＭＳ 明朝"


def make_probe_pure():
    """24 chars: 12 「 + 12 」 alternating."""
    return "「」" * 12  # 24 chars


def make_probe_mixed():
    """24 chars: 12 yak + 12 CJK alternating (control)."""
    return ("「漢」漢" * 6)  # 24 chars: 「漢」漢 × 6 = 24


def make_doc_xml(probe, jc, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
            '<w:sz w:val="24"/></w:rPr>'  # 12pt
            f'<w:t>{probe}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml():
    """Include kern in docDefaults (required for Mech 1)."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault>'
            '<w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
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


def make_docx(label, probe, content_w_pt, jc="both"):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((content_w_pt + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(probe, jc, page_w_tw, margin_tw)
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
                    if t in ("\r","\x07"): continue
                    xs.append((t, float(c.Information(5)), float(c.Information(6)),
                               float(c.Font.Size if c.Font.Size else 0)))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        y0 = xs[0][2]
        line1 = sorted([(t, x, sz) for t, x, y, sz in xs if abs(y - y0) < 0.5],
                       key=lambda v: v[1])
        n_line1 = len(line1)
        char_advances = []
        for i in range(len(line1) - 1):
            t, _, sz = line1[i]
            adv = round(line1[i+1][1] - line1[i][1], 3)
            char_advances.append({"ch": t, "adv": adv, "sz": sz, "ratio": round(adv/sz, 3) if sz > 0 else None})
        return {
            "n_chars_line1": n_line1,
            "char_advances": char_advances,
        }
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}

    # Suite A: pure-yak alternating
    probe_a = make_probe_pure()
    nat_a = 24 * 12.0  # natural = 288pt before Mech 1
    print(f"\n=== Suite A: pure-yak {probe_a!r} natural={nat_a}pt ===")
    out["A_pure"] = {"probe": probe_a, "natural": nat_a, "sweep": []}
    # Sweep around expected post-Mech-1 width (~216pt)
    cw_values_a = [288, 250, 220, 216, 215, 213, 210, 208, 205]
    for cw in cw_values_a:
        try:
            p = make_docx(f"A_pure_cw{cw}", probe_a, cw, "both")
        except Exception as e:
            out["A_pure"]["sweep"].append({"cw": cw, "build_error": str(e)})
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        out["A_pure"]["sweep"].append({"cw": cw, **r})
        # Compute total compression vs natural advance
        if "char_advances" in r:
            total_natural = sum(c["sz"] for c in r["char_advances"])
            actual = sum(c["adv"] for c in r["char_advances"])
            n = r["n_chars_line1"]
            comp = total_natural - actual
            print(f"  cw={cw:>4} n={n} actual={actual:.2f}pt natural_partial={total_natural:.2f} diff={comp:.2f}")
            if r["char_advances"]:
                advs_str = ' '.join(f"{c['ch']}={c['adv']:.1f}" for c in r["char_advances"][:6])
                print(f"     first 6 advs: {advs_str}")
        else:
            print(f"  cw={cw} ERR")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    # Suite B: mixed yak + CJK
    probe_b = make_probe_mixed()
    nat_b = 24 * 12.0  # 288pt
    print(f"\n=== Suite B: mixed {probe_b!r} natural={nat_b}pt ===")
    out["B_mixed"] = {"probe": probe_b, "natural": nat_b, "sweep": []}
    # Mech 1 less aggressive here. natural ~ 240pt after Mech 1 (12 」 compress + 12 漢)
    # Wait: probe = 「漢」漢×6
    # 「(A) preceded by ?: pos 1 not, others preceded by 漢 (CJK) → no compress
    # 」(B) followed by ?: pos 3 followed by 漢 (CJK) → no compress
    # So pure-Mech-1 case: only first 「 compress (followed by 漢)... no:
    #   「 needs prev=A, no Mech 1
    #   」 needs next=A or B, but next is 漢 → no Mech 1
    # All natural width 288pt
    cw_values_b = [290, 285, 282, 280, 278]
    for cw in cw_values_b:
        try:
            p = make_docx(f"B_mixed_cw{cw}", probe_b, cw, "both")
        except Exception as e:
            out["B_mixed"]["sweep"].append({"cw": cw, "build_error": str(e)})
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        out["B_mixed"]["sweep"].append({"cw": cw, **r})
        if "char_advances" in r:
            total_natural = sum(c["sz"] for c in r["char_advances"])
            actual = sum(c["adv"] for c in r["char_advances"])
            n = r["n_chars_line1"]
            comp = total_natural - actual
            print(f"  cw={cw:>4} n={n} actual={actual:.2f} natural_partial={total_natural:.2f} comp={comp:.2f}")
        else:
            print(f"  cw={cw} ERR")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
