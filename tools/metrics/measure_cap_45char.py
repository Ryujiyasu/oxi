"""§4.7b round 18 — Investigate cap discrepancy with 45-char probe.

7f272a P13 L2 (45 chars × MS Mincho 10.5pt × 3 yak) showed total
compression = 7.5pt at slack=7.5, but Round 6/14 formula predicts
cap = floor(sz_val/2)*0.5 = floor(21/2)*0.5 = 5.0pt for fs=10.5.

Hypothesis A: line-length dependent — cap scales with line natural width
Hypothesis B: cap formula context-specific (real doc vs synthetic)
Hypothesis C: P13 had Mech 1 firing I missed
Hypothesis D: Round 6 cap value was specific to 24-char probe length

Test: synthetic probe matching P13 L2 dimensions exactly.
  Probe: 45 chars × MS Mincho 10.5pt × 3 yak at mid-line positions
  Yak: pos 2 (「), pos 9 (」), pos 33 (、) — match P13 L2 positions
  Sweep cw values around expected cap boundary.

If observed cap = 5.0pt, P13 had different conditions (Mech 1?).
If observed cap > 5.0pt, line-length dependence confirmed.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\cap_45char_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\cap_45char.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「」、")
PROBE_LEN = 45
FONT = "ＭＳ 明朝"


def make_probe():
    """45 chars matching P13 L2 yak positions: 「at pos 2, 」at pos 9, 、at pos 33."""
    chars = ["漢"] * PROBE_LEN
    chars[1] = "「"   # pos 2 (1-indexed)
    chars[8] = "」"   # pos 9
    chars[32] = "、"  # pos 33
    return "".join(chars)


def make_doc_xml(probe, jc, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr><w:jc w:val="{jc}"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:hint="eastAsia"/>'
            '<w:sz w:val="21"/></w:rPr>'   # 10.5pt
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


def make_docx(label, probe, content_w_pt, jc="both"):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((content_w_pt + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(probe, jc, page_w_tw, margin_tw)
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
        total_comp = 0.0
        n_yak_comp = 0
        yak_advs = []
        for i in range(len(line1) - 1):
            t, _, sz = line1[i]
            adv = round(line1[i+1][1] - line1[i][1], 3)
            if t in YAKUMONO:
                yak_advs.append({"ch": t, "adv": adv, "sz": sz})
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


def main():
    out = {}
    probe = make_probe()
    natural = PROBE_LEN * 10.5  # 472.5pt
    print(f"probe: {probe!r}")
    print(f"  yak positions: 「pos2, 」pos9, 、pos33")
    print(f"  natural = {natural}pt")
    print(f"  Round 6 formula predicts cap = floor(21/2)*0.5 = 5.0pt for fs=10.5")
    print(f"  P13 L2 observed: total_comp = 7.5pt (over cap by 2.5pt)")

    # Sweep cw values around expected cap boundary
    slacks = [-1, 0, 1, 2, 3, 4, 5, 5.5, 6, 6.5, 7, 7.5, 8, 9, 10, 12]
    cw_values = [round(natural - s, 1) for s in slacks]

    for cw, slack in zip(cw_values, slacks):
        label = f"P_cw{cw:.1f}"
        try:
            p = make_docx(label, probe, cw, "both")
        except Exception as e:
            out[label] = {"build_error": str(e)}
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        out[label] = {"cw": cw, "slack": slack, **r}
        n = r.get("n_chars_line1", "?")
        comp = r.get("total_compression_pt", "?")
        yads = r.get("yak_advs", [])
        yak_str = " ".join(f"{ya['ch']}:{ya['adv']:.1f}" for ya in yads)
        print(f"  cw={cw:>7.1f} slack={slack:>+5.1f} n={n} comp={comp} yak={yak_str}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== SUMMARY ==========")
    sweep = sorted(out.values(), key=lambda x: x.get("slack", 0))
    max_comp_at_full = 0
    first_drop = None
    for r in sweep:
        if r.get("n_chars_line1") == PROBE_LEN:
            tc = r.get("total_compression_pt", 0) or 0
            if tc > max_comp_at_full: max_comp_at_full = tc
        elif first_drop is None and isinstance(r.get("n_chars_line1"), int) and r["n_chars_line1"] < PROBE_LEN and r.get("slack", -1) > 0:
            first_drop = r["slack"]
    print(f"45-char probe at fs=10.5 N=3:")
    print(f"  max_comp_at_full_45_chars: {max_comp_at_full:.2f}pt")
    print(f"  first_drop: {first_drop}")
    print(f"  Round 6 formula predicts: 5.0pt")
    print(f"  P13 L2 observed: 7.5pt")


if __name__ == "__main__":
    main()
