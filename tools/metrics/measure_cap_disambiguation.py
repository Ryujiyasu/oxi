"""§4.7b round 16 — disambiguate cap source.

Round 15 found mixed-CJK case: cap = last_run_fs/2 OR last_yak_fs/2
(indistinguishable since yak fs = surrounding CJK run's fs).

Round 16 forces yak fs ≠ surrounding CJK fs to discriminate:

Suite E1: All-CJK 12pt + last yak forced to fs=14pt
  CJK 24 chars all at 12pt; yak at pos 6/12/18 each in own run.
  pos 6, 12 yak: 12pt; pos 18 yak: 14pt
  - If cap = CJK fs/2 → 6pt
  - If cap = last yak fs/2 → 7pt

Suite E2: CJK 12pt → 14pt (last 12 chars at 14pt) + all yak forced to 12pt
  - last run CJK fs = 14pt, last yak fs = 12pt
  - If cap = last run CJK fs/2 → 7pt
  - If cap = last yak fs/2 → 6pt

Suite E3: All-CJK 12pt + mixed-fs yak (10/12/14pt)
  Test proportional distribution under uniform CJK
  - Total cap = CJK 12/2 = 6pt
  - Distribution: equal (2/2/2)? or proportional (10*6/(10+12+14)≈1.67, 12*6/36=2.0, 14*6/36=2.33)?
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\cap_disambig_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\cap_disambig.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「")
PROBE_LEN = 24
FONT = "ＭＳ 明朝"


def make_runs_e1():
    """All CJK 12pt, last yak (pos 18) at 14pt. Yak at pos 6, 12, 18.
    Returns (text, sz_half) tuple list.
    Pos 6, 12 yak in 12pt context; pos 18 yak forced 14pt."""
    # Layout: CJK1..5 + 「(pos6, 12pt) + CJK7..11 + 「(pos12, 12pt) + CJK13..17 + 「(pos18, 14pt) + CJK19..24
    return [
        ("漢漢漢漢漢", 24),    # 5 chars 12pt
        ("「", 24),            # pos 6: 12pt yak
        ("漢漢漢漢漢", 24),    # 5 chars 12pt
        ("「", 24),            # pos 12: 12pt yak
        ("漢漢漢漢漢", 24),    # 5 chars 12pt
        ("「", 28),            # pos 18: 14pt yak
        ("漢漢漢漢漢漢", 24),  # 6 chars 12pt
    ]


def make_runs_e2():
    """CJK 12pt first 12 + 14pt last 12, all yak forced 12pt."""
    return [
        ("漢漢漢漢漢", 24),    # 5 chars 12pt
        ("「", 24),            # pos 6: 12pt yak
        ("漢漢漢漢漢", 24),    # 5 chars 12pt — first 12 done (12pt)
        ("「", 24),            # pos 12: 12pt yak (overriding 14pt context)
        ("漢漢漢漢漢", 28),    # 5 chars 14pt — start 14pt CJK
        ("「", 24),            # pos 18: 12pt yak (in 14pt CJK section)
        ("漢漢漢漢漢漢", 28),  # 6 chars 14pt
    ]


def make_runs_e3():
    """All CJK 12pt, mixed-fs yak: 10pt, 12pt, 14pt."""
    return [
        ("漢漢漢漢漢", 24),
        ("「", 20),            # pos 6: 10pt yak
        ("漢漢漢漢漢", 24),
        ("「", 24),            # pos 12: 12pt yak
        ("漢漢漢漢漢", 24),
        ("「", 28),            # pos 18: 14pt yak
        ("漢漢漢漢漢漢", 24),
    ]


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
    natural = sum(len(t) * (sz / 2.0) for t, sz in runs)
    print(f"\n=== {label} natural={natural}pt ===")
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
    # Suite E1: CJK 12pt all + last yak forced 14pt
    sweep("E1_uniformCJK_lastyak14", make_runs_e1,
          [-1, 0, 4, 5, 5.5, 6, 6.5, 7, 7.5],
          out, RESULT)
    # Suite E2: Mixed CJK + all yak forced 12pt
    sweep("E2_mixedCJK_yak12", make_runs_e2,
          [-1, 0, 4, 5, 5.5, 6, 6.5, 7, 7.5],
          out, RESULT)
    # Suite E3: All CJK 12pt + mixed yak fs (10/12/14pt)
    sweep("E3_uniformCJK_mixedYak", make_runs_e3,
          [-1, 0, 4, 5, 5.5, 6, 6.5, 7],
          out, RESULT)

    print("\n========== SUMMARY ==========")
    for key in sorted(out.keys()):
        info = out[key]
        sweep_data = sorted(info["sweep"], key=lambda x: x.get("slack", 0))
        max_comp = 0
        first_drop = None
        max_yak_advs = None
        for r in sweep_data:
            n = r.get("n_chars_line1")
            if n == PROBE_LEN:
                tc = r.get("total_compression_pt", 0) or 0
                if tc > max_comp:
                    max_comp = tc
                    max_yak_advs = r.get("yak_advs")
            elif first_drop is None and isinstance(n, int) and n < PROBE_LEN and r.get("slack",-1) > 0:
                first_drop = r["slack"]
        print(f"{key:<30} natural={info['natural']:>6.1f}  max_comp={max_comp:>5.2f}  first_drop={str(first_drop):>5}")
        if max_yak_advs:
            print(f"  at max comp, yak: {max_yak_advs}")


if __name__ == "__main__":
    main()
