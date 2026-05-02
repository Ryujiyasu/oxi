"""§4.7b round 13 — cap formula at 16/18pt + mixed-size line behavior.

Round 6 verified cap = floor(sz_val/2) * 0.5 for 10.5/11/12/14pt MS Mincho.
Extension hypotheses:
  (a) Same formula at 16/18pt: cap = 8pt / 9pt
  (b) Mixed-size line: cap depends on which font size?

Probes (24 chars, N=3 yak mid-line):

Suite A: pure 16pt
  All 24 chars at 16pt MS Mincho, 3 yak at pos 6/12/18
  Expected cap = 16/2 = 8pt

Suite B: pure 18pt
  All 24 chars at 18pt
  Expected cap = 18/2 = 9pt

Suite C: mixed (12pt CJK + 10.5pt yak)
  CJK at 12pt, yak at 10.5pt
  Question: cap = yak's fs/2 (5.25→5.0pt) or dominant fs/2 (6.0pt)?

Suite D: mixed (12pt CJK + 14pt yak)
  CJK at 12pt, yak at 14pt
  Question: cap = yak's 14/2 (7pt) or 12/2 (6pt)?
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\cap_large_mixed_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\cap_large_mixed.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「")
PROBE_LEN = 24
FONT = "ＭＳ 明朝"


def make_probe_n3():
    chars = ["漢"] * PROBE_LEN
    chars[5] = "「"
    chars[11] = "「"
    chars[17] = "「"
    return chars   # return list, caller wraps with sz tags


def make_doc_xml_uniform(probe_chars, font_size_half, jc, page_w_tw, margin_tw):
    """All chars at same size."""
    text = "".join(probe_chars)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{font_size_half}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_doc_xml_mixed(probe_chars, cjk_size_half, yak_size_half, jc, page_w_tw, margin_tw):
    """CJK at one size, yak ('「') at another."""
    runs = []
    cur_text = ""
    cur_size = None
    for ch in probe_chars:
        sz = yak_size_half if ch == "「" else cjk_size_half
        if cur_size is None:
            cur_size = sz
            cur_text = ch
        elif sz == cur_size:
            cur_text += ch
        else:
            # Flush previous run
            runs.append((cur_text, cur_size))
            cur_text = ch
            cur_size = sz
    if cur_text:
        runs.append((cur_text, cur_size))

    runs_xml = ""
    for text, sz in runs:
        runs_xml += (
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r>'
        )

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            f'{runs_xml}'
            '</w:p>'
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


def make_docx(label, doc_xml):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
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
        yak_advs = []
        for i in range(len(line1) - 1):
            t, _, sz = line1[i]
            adv = round(line1[i+1][1] - line1[i][1], 3)
            if t in YAKUMONO:
                n_yak += 1
                yak_advs.append({"adv": adv, "sz": sz, "ratio": round(adv/sz, 3)})
                if sz > 0 and adv < sz * 0.99:
                    n_yak_comp += 1
                    total_comp += (sz - adv)
        return {
            "n_chars_line1": n_line1,
            "total_compression_pt": round(total_comp, 3),
            "n_yak_total_line1": n_yak,
            "n_yak_compressed": n_yak_comp,
            "yak_advs": yak_advs,
        }
    finally:
        try: word.Quit()
        except: pass


def sweep_uniform(label, fs_pt, slacks, out, save_path):
    sz_half = int(round(fs_pt * 2))
    natural = PROBE_LEN * fs_pt
    chars = make_probe_n3()
    print(f"\n=== {label} fs={fs_pt}pt natural={natural}pt cap_theory={fs_pt/2:.1f} ===")
    out[label] = {"font_size_pt": fs_pt, "natural": natural, "cap_theory": fs_pt/2,
                  "kind": "uniform", "sweep": []}
    for slack in slacks:
        cw = round(natural - slack, 1)
        page_w_tw = int((cw + 170) * 20)
        margin_tw = 170 * 10
        doc_xml = make_doc_xml_uniform(chars, sz_half, "both", page_w_tw, margin_tw)
        try:
            p = make_docx(f"{label}_slack{slack}", doc_xml)
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


def sweep_mixed(label, cjk_fs_pt, yak_fs_pt, slacks, out, save_path):
    cjk_half = int(round(cjk_fs_pt * 2))
    yak_half = int(round(yak_fs_pt * 2))
    chars = make_probe_n3()
    n_cjk = sum(1 for c in chars if c == "漢")
    n_yak = sum(1 for c in chars if c == "「")
    natural = n_cjk * cjk_fs_pt + n_yak * yak_fs_pt
    print(f"\n=== {label} CJK={cjk_fs_pt}pt yak={yak_fs_pt}pt natural={natural}pt ===")
    print(f"    Cap hypotheses: yak fs/2={yak_fs_pt/2:.2f}pt vs CJK fs/2={cjk_fs_pt/2:.2f}pt")
    out[label] = {"cjk_fs_pt": cjk_fs_pt, "yak_fs_pt": yak_fs_pt,
                  "natural": natural,
                  "kind": "mixed",
                  "cap_hyp_yak_fs_half": yak_fs_pt/2,
                  "cap_hyp_cjk_fs_half": cjk_fs_pt/2,
                  "sweep": []}
    for slack in slacks:
        cw = round(natural - slack, 1)
        page_w_tw = int((cw + 170) * 20)
        margin_tw = 170 * 10
        doc_xml = make_doc_xml_mixed(chars, cjk_half, yak_half, "both", page_w_tw, margin_tw)
        try:
            p = make_docx(f"{label}_slack{slack}", doc_xml)
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
        yads = r.get("yak_advs", [])
        yak_summary = ", ".join(f"{ya['adv']:.1f}/{ya['sz']:.1f}" for ya in yads)
        print(f"  cw={cw:>6.1f} slack={slack:>+5.1f} n={n} comp={comp}  yak={yak_summary}")
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


def main():
    out = {}

    # Suite A: 16pt
    sweep_uniform("A_fs16", 16.0,
                  [-1, 0, 2, 4, 6, 7, 7.5, 8, 8.5, 9, 10],
                  out, RESULT)
    # Suite B: 18pt
    sweep_uniform("B_fs18", 18.0,
                  [-1, 0, 2, 4, 6, 8, 8.5, 9, 9.5, 10, 11],
                  out, RESULT)
    # Suite C: mixed CJK=12 yak=10.5
    sweep_mixed("C_mixed_cjk12_yak10.5", 12.0, 10.5,
                [-1, 0, 2, 4, 4.5, 5, 5.5, 6, 6.5, 7, 8],
                out, RESULT)
    # Suite D: mixed CJK=12 yak=14
    sweep_mixed("D_mixed_cjk12_yak14", 12.0, 14.0,
                [-1, 0, 2, 4, 5, 5.5, 6, 6.5, 7, 7.5, 8, 9, 10],
                out, RESULT)

    print("\n========== SUMMARY ==========")
    for key in sorted(out.keys()):
        info = out[key]
        sweep = sorted(info["sweep"], key=lambda x: x.get("slack", 0))
        max_comp = 0
        first_drop = None
        for r in sweep:
            n = r.get("n_chars_line1")
            if n == PROBE_LEN:
                tc = r.get("total_compression_pt", 0) or 0
                if tc > max_comp: max_comp = tc
            elif first_drop is None and isinstance(n, int) and n < PROBE_LEN and r.get("slack", -1) > 0:
                first_drop = r["slack"]
        print(f"{key:<32} max_comp={max_comp:>6.2f} first_drop={str(first_drop):>6}")


if __name__ == "__main__":
    main()
