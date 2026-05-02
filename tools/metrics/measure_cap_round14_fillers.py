"""§4.7b round 14 — fs=18 cap fillers + Yu Mincho/Meiryo 16pt cross-font.

Round 13 found:
  fs=18 N=3 MS Mincho: max=8.5pt at slack=8.5; slack=9.0 produced
                       2-char anomaly; slack=9.5 drop. Cap = 8.5..9.0?

Round 14:
  Suite A: fs=18 fillers at slack=8.6, 8.7, 8.8, 8.9, 9.0 (re-test),
           9.1, 9.2 — pin cap exact value for 18pt.
  Suite B: Yu Mincho 16pt N=3 (verify cross-font + cross-size).
  Suite C: Meiryo 16pt N=3 (additional confirmation).
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\cap_round14_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\cap_round14_fillers.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「")
PROBE_LEN = 24


def make_probe_n3():
    chars = ["漢"] * PROBE_LEN
    chars[5] = "「"
    chars[11] = "「"
    chars[17] = "「"
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
        n_yak_comp = 0
        for i in range(len(line1) - 1):
            t, _, sz = line1[i]
            adv = round(line1[i+1][1] - line1[i][1], 3)
            if t in YAKUMONO:
                if sz > 0 and adv < sz * 0.99:
                    n_yak_comp += 1
                    total_comp += (sz - adv)
        return {
            "n_chars_line1": n_line1,
            "total_compression_pt": round(total_comp, 3),
            "n_yak_compressed": n_yak_comp,
        }
    finally:
        try: word.Quit()
        except: pass


def sweep(label, font_name, fs_pt, slacks, out, save_path):
    sz_half = int(round(fs_pt * 2))
    probe = make_probe_n3()
    natural = PROBE_LEN * fs_pt
    cap_th = (sz_half // 2) * 0.5
    print(f"\n=== {label} font={font_name!r} fs={fs_pt}pt natural={natural}pt cap_theory={cap_th} ===")
    out[label] = {"font": font_name, "font_size_pt": fs_pt, "natural": natural,
                  "cap_theory": cap_th, "sweep": []}
    for slack in slacks:
        cw = round(natural - slack, 1)
        sub = f"{label}_slack{slack:.1f}"
        try:
            p = make_docx(sub, probe, cw, font_name, sz_half, "both")
        except Exception as e:
            out[label]["sweep"].append({"slack": slack, "build_error": str(e)})
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
    # Suite A: fs=18 fine fillers around expected cap=9.0
    sweep("A_fs18_filler", "ＭＳ 明朝", 18.0,
          [8.0, 8.5, 8.6, 8.7, 8.8, 8.9, 9.0, 9.1, 9.2, 9.5],
          out, RESULT)
    # Suite B: Yu Mincho 16pt
    sweep("B_YuMincho_fs16", "Yu Mincho", 16.0,
          [-1, 0, 4, 6, 7, 7.5, 8, 8.5, 9],
          out, RESULT)
    # Suite C: Meiryo 16pt
    sweep("C_Meiryo_fs16", "Meiryo", 16.0,
          [-1, 0, 4, 6, 7, 7.5, 8, 8.5, 9],
          out, RESULT)

    print("\n========== SUMMARY ==========")
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
        print(f"{key:<25} {info['font']:<15} fs={info['font_size_pt']:>5} cap_th={info['cap_theory']:>4.1f}  max_comp={max_comp:>5.2f}  first_drop={str(first_drop):>5}")


if __name__ == "__main__":
    main()
