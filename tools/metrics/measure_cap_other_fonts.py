"""§4.7b round 8 — verify cap formula across other CJK fonts.

Round 6/7 confirmed cap = floor(sz_val/2) × 0.5 on MS Mincho. Session 51
found em-dash is font-dependent — possibly cap formula is too.

Test fonts (Japanese installed):
  ＭＳ 明朝 (MS Mincho — baseline)
  Yu Mincho
  Meiryo
  ＭＳ ゴシック (MS Gothic)
  HG ゴシックE

For each: 12pt + 11pt + N=3 mid-line yak + slack sweep.

Plus Suite F: fs=14 N=1 fill to confirm cap=7.0pt (Round 7 had gap).
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\cap_other_fonts_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\cap_other_fonts.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「")
PROBE_LEN = 24

FONTS = [
    "ＭＳ 明朝",      # baseline
    "Yu Mincho",
    "Meiryo",
    "ＭＳ ゴシック",
    "HG明朝E",
]


def make_probe_n3():
    """24-char probe, 3 yak at pos 6, 12, 18 (mid-line)."""
    chars = ["漢"] * PROBE_LEN
    chars[5] = "「"   # pos 6
    chars[11] = "「"  # pos 12
    chars[17] = "「"  # pos 18
    return "".join(chars)


def make_probe_n1():
    chars = ["漢"] * PROBE_LEN
    chars[11] = "「"
    return "".join(chars)


def make_doc_xml(probe, font_name, font_size_half, jc, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}" w:eastAsia="{font_name}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{font_size_half}"/></w:rPr>'
            f'<w:t>{probe}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '</w:settings>')


def make_docx(label, probe, content_w_pt, font_name, font_size_half, jc="both"):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((content_w_pt + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(probe, font_name, font_size_half, jc, page_w_tw, margin_tw)
    settings_xml = make_settings_xml()
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
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
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
        ' Target="settings.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
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
        y0 = xs[0][2]
        line1 = sorted([(t, x, sz) for t, x, y, sz in xs if abs(y - y0) < 0.5],
                       key=lambda v: v[1])
        n_line1 = len(line1)
        total_comp = 0.0
        n_yak = 0
        n_yak_comp = 0
        for i in range(len(line1) - 1):
            t, _, sz = line1[i]
            adv = round(line1[i+1][1] - line1[i][1], 3)
            if t in YAKUMONO:
                n_yak += 1
                if sz > 0 and adv < sz * 0.99:
                    n_yak_comp += 1
                    total_comp += (sz - adv)
        return {
            "n_chars_line1": n_line1,
            "total_compression_pt": round(total_comp, 3),
            "n_yak_total_line1": n_yak,
            "n_yak_compressed": n_yak_comp,
        }
    finally:
        try: word.Quit()
        except: pass


def sweep(label, font_name, font_size_pt, n_yak, probe_fn, out, save_path):
    font_size_half = int(round(font_size_pt * 2))
    probe = probe_fn()
    natural = PROBE_LEN * font_size_pt
    cap_theory = (font_size_half // 2) * 0.5  # = floor(fs)*0.5 in 0.5pt steps
    test_slacks = [-1, 0, 1, 2, 3, cap_theory-1, cap_theory-0.5, cap_theory,
                   cap_theory+0.5, cap_theory+1, cap_theory+2, cap_theory+5,
                   font_size_pt+1]
    test_slacks = sorted(set(round(s, 1) for s in test_slacks if s >= -1))
    cw_values = [round(natural - s, 1) for s in test_slacks]
    print(f"\n=== {label} font={font_name!r} fs={font_size_pt}pt N={n_yak} cap_theory={cap_theory:.1f} ===")
    out[label] = {
        "font": font_name, "font_size_pt": font_size_pt, "n_yak": n_yak,
        "cap_theory": cap_theory, "natural": natural, "sweep": [],
    }
    for cw, slack in zip(cw_values, test_slacks):
        sub = f"{label}_cw{cw:.1f}"
        try:
            p = make_docx(sub, probe, cw, font_name, font_size_half, "both")
        except Exception as e:
            out[label]["sweep"].append({"cw": cw, "slack": slack, "build_error": str(e)})
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        out[label]["sweep"].append({"cw": cw, "slack": slack, **r})
        n = r.get("n_chars_line1", "?")
        comp = r.get("total_compression_pt", "?")
        print(f"  cw={cw:>6.1f} slack={slack:>+5.1f} n_line1={n} total_comp={comp}")
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


def main():
    out = {}
    # Suite E: 5 fonts × 12pt × N=3
    print("\n========== Suite E: 5 fonts × 12pt × N=3 ==========")
    for fn in FONTS:
        font_id = fn.replace(" ", "").replace("ＭＳ", "MS").replace("明朝", "Min").replace("ゴシック", "Got")
        sweep(f"E_{font_id}_fs12_N3", fn, 12.0, 3, make_probe_n3, out, RESULT)

    # Suite F: fs=14 N=1 gap-fill (fill slack 7..14)
    print("\n========== Suite F: fs=14 N=1 gap-fill ==========")
    # Use MS Mincho, N=1, slack values 7, 8, 9, 10, 11, 12, 13, 14
    label = "F_MSMin_fs14_N1_gapfill"
    out[label] = {
        "font": "ＭＳ 明朝", "font_size_pt": 14.0, "n_yak": 1,
        "cap_theory": 7.0, "natural": 24*14, "sweep": [],
    }
    probe = make_probe_n1()
    for slack in [7, 8, 9, 10, 11, 12, 13, 14]:
        cw = 24 * 14 - slack
        sub = f"{label}_slack{slack}"
        try:
            p = make_docx(sub, probe, cw, "ＭＳ 明朝", 28, "both")
        except Exception as e:
            out[label]["sweep"].append({"cw": cw, "slack": slack, "build_error": str(e)})
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        out[label]["sweep"].append({"cw": cw, "slack": slack, **r})
        n = r.get("n_chars_line1", "?")
        comp = r.get("total_compression_pt", "?")
        print(f"  cw={cw} slack={slack} n_line1={n} total_comp={comp}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    # Summary
    print("\n========== SUMMARY ==========")
    print(f"{'label':<35} {'font':<18} {'fs':>4} {'N':>3} {'cap_th':>7} {'max_comp':>9} {'first_drop':>11}")
    for key in sorted(out.keys()):
        info = out[key]
        sweep_data = sorted(info["sweep"], key=lambda x: x.get("slack", 0))
        max_comp = 0
        first_drop = None
        for r in sweep_data:
            n = r.get("n_chars_line1")
            if n == PROBE_LEN:
                tc = r.get("total_compression_pt", 0) or 0
                if tc > max_comp: max_comp = tc
            elif first_drop is None and isinstance(n, int) and n < PROBE_LEN and r.get("slack", -1) > 0:
                first_drop = r["slack"]
        print(f"{key:<35} {info['font'][:17]:<18} {info['font_size_pt']:>4} {info['n_yak']:>3} {info['cap_theory']:>7.1f} {max_comp:>9.2f} {str(first_drop):>11}")


if __name__ == "__main__":
    main()
