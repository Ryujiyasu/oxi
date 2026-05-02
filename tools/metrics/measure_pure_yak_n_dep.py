"""§4.7b round 21 — pure-yak Mech 2 cap n_yak dependence + fs=10.5 boundary.

Round 20 produced cap = 0.75 × fs at fs ∈ {11,12,14,16} (4/5 confirm)
with single probe length n_compressible=11.

Round 21 questions:
  Q1: fs=10.5 fine sweep — predicted cap = 7.5pt (= floor(15.75)*0.5).
      Probe slack ∈ {6.0, 6.5, 7.0, 7.5, 7.875, 8.0, 8.5}.
  Q2: cap dependence on n_compressible — vary probe length to give
      different counts of mid-line 「 chars at fixed fs=12pt.
      Length ∈ {12, 16, 20, 24, 30, 40}.
      n_compressible = (length-2)/2 (excluding line-start/end).
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\pure_yak_n_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\pure_yak_n.json")
os.makedirs(OUT_DIR, exist_ok=True)

FONT = "ＭＳ 明朝"


def make_doc_xml(probe, jc, page_w_tw, margin_tw, sz_val):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{probe}</w:t></w:r></w:p>'
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
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
            '<w:kern w:val="2"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '</w:settings>')


def make_docx(label, probe, content_w_pt, sz_val):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((content_w_pt + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(probe, "both", page_w_tw, margin_tw, sz_val)
    styles_xml = make_styles_xml(sz_val)
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
            char_advances.append({"ch": t, "adv": adv, "sz": sz})
        return {
            "n_chars_line1": n_line1,
            "char_advances": char_advances,
            "advs_first_8": " ".join(f"{c['ch']}={c['adv']:.1f}" for c in char_advances[:8]),
        }
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}

    # Q1: fs=10.5 fine sweep
    print("=== Q1: fs=10.5 fine sweep ===")
    fs_pt = 10.5
    sz_val = 21
    nat = 18.0 * fs_pt  # 189.0pt
    probe = "「」" * 12  # 24 chars
    out["Q1_fs10p5"] = {"fs": fs_pt, "n_compressible": 11, "sweep": []}
    for slack in [6.0, 6.5, 7.0, 7.5, 7.875, 8.0, 8.5]:
        cw = round(nat - slack, 3)
        label = f"Q1_sl{slack}"
        try:
            p = make_docx(label, probe, cw, sz_val)
        except Exception as e:
            out["Q1_fs10p5"]["sweep"].append({"slack": slack, "cw": cw, "build_error": str(e)})
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        entry = {"slack": slack, "cw": cw, **r}
        advs = entry.pop("char_advances", None)
        if advs:
            sum_actual = sum(c["adv"] for c in advs)
            sum_nat_post_mech1 = sum(
                fs_pt if c["ch"] == "「" else fs_pt/2 for c in advs)
            entry["mech2_comp_pt"] = round(sum_nat_post_mech1 - sum_actual, 2)
        out["Q1_fs10p5"]["sweep"].append(entry)
        n = entry.get("n_chars_line1", "?")
        adv_str = entry.get("advs_first_8", "")[:60]
        m2 = entry.get("mech2_comp_pt", "?")
        print(f"  slack={slack:>5.3f} cw={cw:>7.2f} n={n} m2={m2} | {adv_str}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    # Q2: n_compressible dependence at fs=12.0
    print("\n=== Q2: n_compressible sweep at fs=12 ===")
    fs_pt = 12.0
    sz_val = 24
    out["Q2_n_dep"] = {"fs": fs_pt, "tests": []}
    # Length L → n_yak (Type A 「) compressible mid-line:
    #   L=12 (24 in alternate) → n_compressible 「 = 5
    #   L=16 → 7
    #   L=20 → 9
    #   L=24 → 11
    #   L=30 → 14
    # Mech 1 reduces 」 to half. nat after Mech 1 = (L/2) × 12 + (L/2) × 6 = 9L
    # Cap candidates: if n_dep, expect cap ∝ n_compressible.
    #   If cap=0.75×fs constant, all give cap=9pt regardless of L.
    # Probe each L at slacks {3, 6, 9, 12} (= 0.25/0.5/0.75/1.0 × fs).
    for L in [12, 16, 20, 24, 30]:
        # L must be even (alternating 「」)
        probe = "「」" * (L // 2)
        n_compressible = (L - 2) // 2  # mid-line 「
        nat_post_mech1 = 9.0 * L  # 12*L/2 + 6*L/2
        for slack in [3.0, 6.0, 9.0, 12.0]:
            cw = round(nat_post_mech1 - slack, 3)
            label = f"Q2_L{L}_sl{slack}"
            try:
                p = make_docx(label, probe, cw, sz_val)
            except Exception as e:
                out["Q2_n_dep"]["tests"].append({"L": L, "n_compressible": n_compressible,
                                                  "slack": slack, "cw": cw, "build_error": str(e)})
                continue
            kill_word()
            try:
                r = measure_one(p)
            except Exception as e:
                r = {"measure_error": str(e)}
                kill_word()
            entry = {"L": L, "n_compressible": n_compressible, "slack": slack, "cw": cw, **r}
            advs = entry.pop("char_advances", None)
            if advs:
                sum_actual = sum(c["adv"] for c in advs)
                sum_nat_post_mech1 = sum(
                    fs_pt if c["ch"] == "「" else fs_pt/2 for c in advs)
                entry["mech2_comp_pt"] = round(sum_nat_post_mech1 - sum_actual, 2)
            out["Q2_n_dep"]["tests"].append(entry)
            n = entry.get("n_chars_line1", "?")
            m2 = entry.get("mech2_comp_pt", "?")
            print(f"  L={L:>3} n_comp={n_compressible:>2} slack={slack:>4.1f} cw={cw:>6.1f} n={n} m2={m2}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== Q2 SUMMARY ==========")
    print(f"{'L':>4} {'n_comp':>7} {'slack=3':>10} {'slack=6':>10} {'slack=9':>10} {'slack=12':>10}")
    for L in [12, 16, 20, 24, 30]:
        row = {"L": L}
        for t in out["Q2_n_dep"]["tests"]:
            if t.get("L") == L:
                row[t["slack"]] = t.get("n_chars_line1", "?")
        print(f"  {L:>3} {(L-2)//2:>5} {row.get(3.0,'?'):>10} {row.get(6.0,'?'):>10} "
              f"{row.get(9.0,'?'):>10} {row.get(12.0,'?'):>10}")
    print("\n  fits=L means Mech 2 cap covered the slack")
    print("  fit drops at slack > cap → cap ≈ slack just before drop")


if __name__ == "__main__":
    main()
