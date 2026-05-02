"""§4.7b round 20 — pure-yak Mech 2 cap formula via fs sweep.

Round 19 confirmed Mech 2 fires on pure-yak lines (no CJK).
Round 20 derives the cap formula by sweeping fs and counting
mid-line compressible 「 chars.

Probe template (pure-yak alternating「」, 24 chars):
  After Mech 1 fires alternating: 」=half (sz_val_int/2 × 0.5 pt),
  「 retained at full sz_val × 0.5 pt.
  Line natural after Mech 1:
    = 12 × sz + 12 × (sz/2)  for fs in pt (where sz = fs)
    = 12 × fs + 12 × (fs/2) = 18 × fs
  Compressible chars = 11 mid-line 「 (pos 3, 5, ..., 23)
  pos 1 line-start exempt, pos 24 (」) line-end exempt.

Find max slack (= cap_total) at which n_chars_line1 still equals 24.

Test fs values: 10.5, 11.0, 12.0, 14.0, 16.0
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\pure_yak_fs_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\pure_yak_fs.json")
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
        # Build advances (only first 10 for diagnostic display; full sum for total)
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
    probe = "「」" * 12  # 24 chars
    print(f"Probe: pure-yak alternating, 24 chars: {probe!r}")
    print(f"Expected after Mech 1: 12 「 (full) + 12 」 (half) = 18 × fs pt")
    print(f"Compressible mid-line 「 chars = 11 (positions 3..23 odd)\n")

    fs_values = [10.5, 11.0, 12.0, 14.0, 16.0]
    for fs_pt in fs_values:
        sz_val = int(round(fs_pt * 2))
        nat_after_mech1 = 18.0 * fs_pt
        print(f"=== fs={fs_pt}pt (sz val={sz_val}) nat_after_mech1={nat_after_mech1}pt ===")
        out[f"fs{fs_pt}"] = {"fs": fs_pt, "sz_val": sz_val, "nat_after_mech1": nat_after_mech1, "sweep": []}
        # Probe key slacks only: bracket cap via 7 points
        # Hypothesis: cap ~ fs/2 to fs (test 0.5×fs, 0.75×fs, 1.0×fs, 1.25×fs)
        slack_steps = sorted({
            round(fs_pt * 0.25, 1),
            round(fs_pt * 0.5, 1),
            round(fs_pt * 0.75, 1),
            round(fs_pt * 1.0, 1),
            round(fs_pt * 1.25, 1),
            round(fs_pt * 1.5, 1),
            round(fs_pt * 2.0, 1),
        })
        for slack in slack_steps:
            cw = round(nat_after_mech1 - slack, 2)
            label = f"fs{fs_pt}_sl{slack}"
            try:
                p = make_docx(label, probe, cw, sz_val)
            except Exception as e:
                out[f"fs{fs_pt}"]["sweep"].append({"slack": slack, "cw": cw, "build_error": str(e)})
                continue
            kill_word()
            try:
                r = measure_one(p)
            except Exception as e:
                r = {"measure_error": str(e)}
                kill_word()
            entry = {"slack": slack, "cw": cw, **r}
            # drop full advances list for compactness
            advs = entry.pop("char_advances", None)
            if advs:
                # Compute Mech 2 actual savings: sum(natural advs) - sum(actual)
                # Natural after Mech 1: 「=fs, 」=fs/2
                # We have first 23 advances. Sum natural = sum_{first23} (fs if 「 else fs/2)
                # If line broke, first n-1 advances are within line.
                first_n_minus_1 = advs
                sum_actual = sum(c["adv"] for c in first_n_minus_1)
                sum_natural_post_mech1 = sum(
                    fs_pt if c["ch"] == "「" else (fs_pt / 2.0)
                    for c in first_n_minus_1
                )
                entry["sum_actual_adv"] = round(sum_actual, 2)
                entry["sum_nat_post_mech1"] = round(sum_natural_post_mech1, 2)
                entry["mech2_compression_pt"] = round(sum_natural_post_mech1 - sum_actual, 2)
            out[f"fs{fs_pt}"]["sweep"].append(entry)
            n = entry.get("n_chars_line1", "?")
            mech2c = entry.get("mech2_compression_pt", "?")
            advs_str = entry.get("advs_first_8", "")[:60]
            print(f"  slack={slack:>4.1f} cw={cw:>7.2f} n={n} mech2_comp={mech2c} | {advs_str}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)
        # Find cap_max (last slack where n=24)
        passing = [e for e in out[f"fs{fs_pt}"]["sweep"] if e.get("n_chars_line1") == 24]
        if passing:
            cap_max = max(e["slack"] for e in passing)
            out[f"fs{fs_pt}"]["cap_max_slack_pt"] = cap_max
            print(f"  >> cap_max_slack at n=24: {cap_max}pt (= Mech 2 absolute cap)\n")

    print("\n========== SUMMARY ==========")
    print(f"{'fs(pt)':>8} {'sz_val':>8} {'cap_max_pt':>12} {'fs/2':>6} {'cap/fs':>8}")
    for fs_pt in fs_values:
        info = out.get(f"fs{fs_pt}", {})
        cm = info.get("cap_max_slack_pt", "?")
        if isinstance(cm, (int, float)):
            print(f"{fs_pt:>8} {info['sz_val']:>8} {cm:>12.2f} {fs_pt/2:>6.2f} {cm/fs_pt:>8.3f}")


if __name__ == "__main__":
    main()
