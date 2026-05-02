"""§4.7b round 10 — N=1 cap sweep gap fillers.

Round 7 had sweep gaps:
  fs=10.5: tested up to slack=4.5 then jumped to 5.5 (drop)
  fs=11.0: tested up to slack=4.7 then jumped to 5.7 (drop)
  fs=12.0: tested up to slack=6.0 then jumped to 12.5 (drop)
  fs=14.0: tested up to slack=6.7 then jumped to 14.5 (drop)

Round 8 fs=14 N=1 gap-fill confirmed first_drop=8.0 (cap+1).
This round fills remaining 3 fonts at the cap boundary.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\n1_cap_filler_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\n1_cap_fillers.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「")
PROBE_LEN = 24
FONT = "ＭＳ 明朝"


def make_probe_n1():
    chars = ["漢"] * PROBE_LEN
    chars[11] = "「"  # pos 12 mid-line
    return "".join(chars)


def make_doc_xml(probe, font_size_half, jc, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
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


def make_docx(label, probe, content_w_pt, font_size_half, jc="both"):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((content_w_pt + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(probe, font_size_half, jc, page_w_tw, margin_tw)
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


# Filler slack values per font (focus on gap regions)
FILL_PLAN = {
    10.5: [4.6, 4.8, 5.0, 5.1, 5.2, 5.3, 5.4],   # fill 4.5..5.5
    11.0: [4.8, 5.0, 5.2, 5.4, 5.5, 5.6],          # fill 4.7..5.7
    12.0: [6.5, 7.0, 7.5, 8.0, 9.0, 10.0, 12.0, 12.4],  # fill 6.0..12.5
    14.0: [6.8, 7.0, 7.2, 7.5, 7.8],               # fill 6.7..8.0 (R8 confirmed)
}


def main():
    out = {}
    probe = make_probe_n1()
    print(f"probe: {probe!r} (yak at pos 12)")
    for fs in sorted(FILL_PLAN.keys()):
        sz_half = int(round(fs * 2))
        natural = PROBE_LEN * fs
        cap_theory = (sz_half // 2) * 0.5
        slacks = FILL_PLAN[fs]
        print(f"\n=== fs={fs}pt natural={natural}pt cap_theory={cap_theory}pt ===")
        key = f"fs{fs}_N1"
        out[key] = {
            "font_size_pt": fs, "n_yak": 1, "natural": natural,
            "cap_theory": cap_theory, "sweep": [],
        }
        for slack in slacks:
            cw = round(natural - slack, 1)
            label = f"{key}_slack{slack:.1f}"
            try:
                p = make_docx(label, probe, cw, sz_half, "both")
            except Exception as e:
                out[key]["sweep"].append({"slack": slack, "build_error": str(e)})
                continue
            kill_word()
            try:
                r = measure_one(p)
            except Exception as e:
                r = {"measure_error": str(e)}
                kill_word()
            out[key]["sweep"].append({"cw": cw, "slack": slack, **r})
            n = r.get("n_chars_line1", "?")
            comp = r.get("total_compression_pt", "?")
            print(f"  cw={cw:>6.1f} slack={slack:>+5.1f} n_line1={n} total_comp={comp}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== SUMMARY ==========")
    for fs in sorted(FILL_PLAN.keys()):
        key = f"fs{fs}_N1"
        info = out[key]
        sweep = info["sweep"]
        max_comp = 0
        first_drop = None
        for r in sorted(sweep, key=lambda x: x.get("slack", 0)):
            n = r.get("n_chars_line1")
            if n == PROBE_LEN:
                tc = r.get("total_compression_pt", 0) or 0
                if tc > max_comp: max_comp = tc
            elif first_drop is None and isinstance(n, int) and n < PROBE_LEN and r.get("slack", -1) > 0:
                first_drop = r["slack"]
        print(f"fs={fs}pt cap_theory={info['cap_theory']:.2f} max_comp_observed={max_comp:.2f} first_drop={first_drop}")


if __name__ == "__main__":
    main()
