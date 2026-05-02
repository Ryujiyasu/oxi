"""§4.7b round 15 — Mixed CJK sizes within one line.

Round 13 found cap = CJK fs/2 in mixed-yak-size case (yak at different
fs but CJK uniform). Round 15 tests when CJK itself has multiple sizes.

Probe (24 chars): 12pt CJK first half + 14pt CJK second half + 3 yak
distributed.

Hypotheses for cap:
  (a) Max CJK fs: cap = 14/2 = 7pt
  (b) Min CJK fs: cap = 12/2 = 6pt
  (c) First-run CJK fs: cap = 12/2 = 6pt
  (d) Last-run CJK fs: cap = 14/2 = 7pt
  (e) Per-yak-run CJK fs (yak's own surrounding context)
  (f) Avg/proportional

Test variants:
  Suite A (baseline): all 12pt → cap = 6pt
  Suite B: 12pt first 12 chars + 14pt last 12 chars; yak at pos 6, 12, 18
  Suite C: 14pt first 12 + 12pt last 12; yak at pos 6, 12, 18 (reverse of B)

If (a) max wins: B=7, C=7
If (b) min wins: B=6, C=6
If (c) first-run wins: B=6, C=7
If (d) last-run wins: B=7, C=6
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\mixed_cjk_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\mixed_cjk_sizes.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「")
PROBE_LEN = 24
FONT = "ＭＳ 明朝"


def make_probe_runs_a():
    """Suite A: all 12pt baseline. Returns runs = [(text, sz_half)]."""
    chars = ["漢"] * 24
    chars[5] = "「"; chars[11] = "「"; chars[17] = "「"
    return [("".join(chars), 24)]   # all 12pt


def make_probe_runs_b():
    """Suite B: first 12 chars 12pt, last 12 chars 14pt. Yak in mixed positions."""
    chars1 = ["漢"] * 12
    chars1[5] = "「"; chars1[11] = "「"   # yak at pos 6, 12 (in 12pt run)
    chars2 = ["漢"] * 12
    chars2[5] = "「"   # yak at pos 18 (= run2 pos 6, in 14pt run)
    return [("".join(chars1), 24), ("".join(chars2), 28)]


def make_probe_runs_c():
    """Suite C: first 12 chars 14pt, last 12 chars 12pt. (Reverse of B)"""
    chars1 = ["漢"] * 12
    chars1[5] = "「"; chars1[11] = "「"   # yak in 14pt run
    chars2 = ["漢"] * 12
    chars2[5] = "「"   # yak in 12pt run
    return [("".join(chars1), 28), ("".join(chars2), 24)]


def make_doc_xml(runs, jc, page_w_tw, margin_tw):
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


def make_docx(label, runs, content_w_pt, jc="both"):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((content_w_pt + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(runs, jc, page_w_tw, margin_tw)
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
        yak_advs = []
        for i in range(len(line1) - 1):
            t, _, sz = line1[i]
            adv = round(line1[i+1][1] - line1[i][1], 3)
            if t in YAKUMONO:
                yak_advs.append({"adv": adv, "sz": sz})
                if sz > 0 and adv < sz * 0.99:
                    n_yak_comp += 1
                    total_comp += (sz - adv)
        return {
            "n_chars_line1": n_line1,
            "total_compression_pt": round(total_comp, 3),
            "n_yak_compressed": n_yak_comp,
            "yak_advs": yak_advs,
        }
    finally:
        try: word.Quit()
        except: pass


def sweep(label, runs_fn, slacks, out, save_path):
    runs = runs_fn()
    # Compute natural width
    natural = sum(len(t) * (sz / 2.0) for t, sz in runs)
    print(f"\n=== {label} natural={natural}pt runs={[(len(t), sz/2) for t,sz in runs]} ===")
    out[label] = {"natural": natural, "runs": [(t, sz) for t, sz in runs], "sweep": []}
    for slack in slacks:
        cw = round(natural - slack, 1)
        sub = f"{label}_slack{slack:.1f}"
        try:
            p = make_docx(sub, runs, cw, "both")
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
        yak_str = " ".join(f"{ya['adv']:.1f}/{ya['sz']:.1f}" for ya in yads)
        print(f"  cw={cw:>6.1f} slack={slack:>+5.1f} n={n} comp={comp}  yak={yak_str}")
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


def main():
    out = {}
    # Suite A: pure 12pt (baseline cap=6pt)
    sweep("A_pure_12pt", make_probe_runs_a,
          [-1, 0, 4, 6, 6.5, 7, 8],
          out, RESULT)
    # Suite B: 12pt first half + 14pt second half
    sweep("B_12then14", make_probe_runs_b,
          [-1, 0, 4, 6, 6.5, 7, 7.5, 8],
          out, RESULT)
    # Suite C: 14pt first half + 12pt second half (reverse of B)
    sweep("C_14then12", make_probe_runs_c,
          [-1, 0, 4, 6, 6.5, 7, 7.5, 8],
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
        print(f"{key:<20} natural={info['natural']:>6.1f}  max_comp={max_comp:>5.2f}  first_drop={str(first_drop):>5}")


if __name__ == "__main__":
    main()
