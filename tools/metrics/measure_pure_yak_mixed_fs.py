"""§4.7b round 22 — mixed-fs pure-yak (which run drives cap).

Pure-yak line with two yak runs at different sizes. Round 16 mixed-line
formula uses the LAST CJK run's fs. Does pure-yak follow the same rule?

Probe: 24 chars, alternating 「」, two runs of 12 chars each.

Suite A: first half 12pt + second half 14pt
Suite B: first half 14pt + second half 12pt (reversed)
Suite C: first half 12pt + second half 16pt (larger gap)

Predicted caps:
  H1 (first run drives): A→9.0, B→10.5, C→9.0
  H2 (last run drives):  A→10.5, B→9.0, C→12.0
  H3 (max run drives):   A→10.5, B→10.5, C→12.0
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\pure_yak_mixed_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\pure_yak_mixed.json")
os.makedirs(OUT_DIR, exist_ok=True)

FONT = "ＭＳ 明朝"


def make_doc_xml(probe_a, probe_b, sz_a, sz_b, jc, page_w_tw, margin_tw):
    """Two runs in one paragraph, different sizes."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_a}"/></w:rPr>'
            f'<w:t>{probe_a}</w:t></w:r>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_b}"/></w:rPr>'
            f'<w:t>{probe_b}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml(default_sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
            '<w:kern w:val="2"/>'
            f'<w:sz w:val="{default_sz}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '</w:settings>')


def make_docx(label, probe_a, probe_b, sz_a, sz_b, content_w_pt):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((content_w_pt + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(probe_a, probe_b, sz_a, sz_b, "both", page_w_tw, margin_tw)
    styles_xml = make_styles_xml(sz_a)
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
            "advs_first_8": " ".join(f"{c['ch']}({c['sz']:.0f})={c['adv']:.1f}" for c in char_advances[:8]),
            "advs_around_split": " ".join(f"{c['ch']}({c['sz']:.0f})={c['adv']:.1f}" for c in char_advances[10:14]),
            "advs_last_4": " ".join(f"{c['ch']}({c['sz']:.0f})={c['adv']:.1f}" for c in char_advances[-4:]),
        }
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    half_probe = "「」" * 6  # 12 chars

    # Three suites: each tests slack at predicted cap_first vs cap_last
    suites = [
        # (label, sz_a, sz_b, predict_first, predict_last, predict_max)
        ("A_12_14",  24, 28,  9.0, 10.5, 10.5),
        ("B_14_12",  28, 24, 10.5,  9.0, 10.5),
        ("C_12_16",  24, 32,  9.0, 12.0, 12.0),
        ("D_16_12",  32, 24, 12.0,  9.0, 12.0),
    ]

    for label, sz_a, sz_b, p_first, p_last, p_max in suites:
        fs_a = sz_a / 2
        fs_b = sz_b / 2
        # nat after Mech 1: half A: 12 chars (6 「 + 6 」) → 6×fs_a + 6×fs_a/2 = 9×fs_a
        # half B: 9×fs_b
        nat_after_mech1 = 9 * fs_a + 9 * fs_b
        print(f"\n=== {label}: A={fs_a}pt, B={fs_b}pt nat_after_mech1={nat_after_mech1}pt ===")
        print(f"  predict_first={p_first}, predict_last={p_last}, predict_max={p_max}")
        out[label] = {"sz_a": sz_a, "sz_b": sz_b, "fs_a": fs_a, "fs_b": fs_b,
                      "nat_after_mech1": nat_after_mech1,
                      "predict_first": p_first, "predict_last": p_last, "predict_max": p_max,
                      "sweep": []}
        # Test slacks bracketing all three predictions
        slack_steps = sorted(set([p_first - 0.5, p_first + 0.5,
                                  p_last - 0.5, p_last + 0.5,
                                  p_max - 0.5, p_max + 0.5]))
        slack_steps = [s for s in slack_steps if s > 0]
        for slack in slack_steps:
            cw = round(nat_after_mech1 - slack, 3)
            sublabel = f"{label}_sl{slack}"
            try:
                p = make_docx(sublabel, half_probe, half_probe, sz_a, sz_b, cw)
            except Exception as e:
                out[label]["sweep"].append({"slack": slack, "cw": cw, "build_error": str(e)})
                continue
            kill_word()
            try:
                r = measure_one(p)
            except Exception as e:
                r = {"measure_error": str(e)}
                kill_word()
            entry = {"slack": slack, "cw": cw, **r}
            out[label]["sweep"].append(entry)
            n = entry.get("n_chars_line1", "?")
            f8 = entry.get("advs_first_8", "")[:60]
            print(f"  slack={slack:>5.1f} cw={cw:>7.2f} n={n} | {f8}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)
        # Determine cap_max
        passing = [e for e in out[label]["sweep"] if e.get("n_chars_line1") == 24]
        if passing:
            cap_max = max(e["slack"] for e in passing)
            out[label]["cap_max_observed"] = cap_max
            print(f"  >> cap_max = {cap_max}pt")
            # Match against predictions
            matches = []
            if abs(cap_max - p_first) <= 0.5: matches.append("first")
            if abs(cap_max - p_last) <= 0.5: matches.append("last")
            if abs(cap_max - p_max) <= 0.5: matches.append("max")
            print(f"  >> matches: {matches}")

    print("\n========== SUMMARY ==========")
    print(f"{'suite':>10} {'first':>7} {'last':>7} {'max':>7} {'observed':>10} {'matches':>15}")
    for label, sz_a, sz_b, p_first, p_last, p_max in suites:
        info = out.get(label, {})
        cm = info.get("cap_max_observed", "?")
        cm_str = f"{cm:.1f}" if isinstance(cm, (int, float)) else cm
        matches = []
        if isinstance(cm, (int, float)):
            if abs(cm - p_first) <= 0.5: matches.append("first")
            if abs(cm - p_last) <= 0.5: matches.append("last")
            if abs(cm - p_max) <= 0.5: matches.append("max")
        print(f"{label:>10} {p_first:>7.1f} {p_last:>7.1f} {p_max:>7.1f} {cm_str:>10} {','.join(matches):>15}")


if __name__ == "__main__":
    main()
