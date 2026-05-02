"""§4.7b round 6 — verify line-level cap formula across font sizes.

Round 5 found cap = fontSize/2 at 12pt for N ∈ {3, 4, 5, 7}.
This script verifies universality across 10.5pt / 11pt / 12pt / 14pt.

Tests:
  Suite A: font_size × {10.5, 11, 12, 14} × N=3 yak × cw sweep
  Suite B: 12pt × N=2 with MID-LINE yak placement (positions 8, 16)
           (resolves Round 5's line-end yak anomaly)

Probe: 24-char line, N yak evenly spaced in mid-line (NOT at pos 1, NOT
at pos 24).
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\cap_font_sweep_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\cap_font_sweep.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「")
PROBE_LEN = 24
FONT = "ＭＳ 明朝"


def make_probe(n_yak: int, total: int = PROBE_LEN) -> str:
    """24-char probe with n_yak yakumono at mid-line positions
    (positions 2..total-2 inclusive, evenly spaced)."""
    chars = ["漢"] * total
    if n_yak == 0: return "".join(chars)
    # Use safe range (avoid pos 1 line-start and pos total line-end)
    # For total=24, valid yak positions = 2..22 (inclusive, 1-indexed)
    # Positions in 0-indexed = 1..21
    # Distribute n_yak evenly in this range
    safe_first = 1
    safe_last = total - 3  # 0-indexed pos 21 (1-indexed 22)
    if n_yak == 1:
        positions = [(safe_first + safe_last) // 2]
    else:
        step = (safe_last - safe_first) / (n_yak - 1)
        positions = [int(safe_first + i * step) for i in range(n_yak)]
    for p in positions:
        chars[p] = "「"
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
        d = word.Documents.Open(path, ReadOnly=True)
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
        n_yak_total = 0
        n_yak_compressed = 0
        for i in range(len(line1) - 1):
            t = line1[i][0]
            adv = round(line1[i+1][1] - line1[i][1], 3)
            sz = line1[i][2]
            if t in YAKUMONO:
                n_yak_total += 1
                if sz > 0 and adv < sz * 0.99:
                    n_yak_compressed += 1
                    total_comp += (sz - adv)
        return {
            "n_chars_line1": n_line1,
            "total_compression_pt": round(total_comp, 3),
            "n_yak_total_line1": n_yak_total,
            "n_yak_compressed": n_yak_compressed,
        }
    finally:
        try: word.Quit()
        except: pass


def sweep_for_drop(font_size_pt, n_yak, suite_label):
    font_size_half = int(round(font_size_pt * 2))
    probe = make_probe(n_yak)
    natural = PROBE_LEN * font_size_pt
    expected_cap = font_size_pt / 2.0
    # Sweep cw: focus around expected drop boundary (slack near cap+1)
    slacks = [-1, 0, 1, 2, 3, expected_cap-1, expected_cap, expected_cap+1, expected_cap+2, expected_cap+3, expected_cap+5, font_size_pt+1, font_size_pt*2]
    slacks = sorted(set([round(s, 1) for s in slacks if s >= -1]))
    cw_values = [round(natural - s, 1) for s in slacks]
    print(f"\n=== {suite_label} fontSize={font_size_pt}pt N={n_yak} natural={natural}pt expected_cap={expected_cap}pt ===")
    print(f"  positions of yak in probe: {[i+1 for i, c in enumerate(probe) if c in YAKUMONO]}")
    results = []
    for cw, slack in zip(cw_values, slacks):
        label = f"{suite_label}_fs{font_size_pt}_N{n_yak}_cw{cw:.1f}"
        try:
            p = make_docx(label, probe, cw, font_size_half, "both")
        except Exception as e:
            results.append({"cw": cw, "slack": slack, "build_error": str(e)})
            continue
        kill_word()
        try:
            r = measure_one(p)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        results.append({"cw": cw, "slack": slack, **r})
        n = r.get("n_chars_line1", "?")
        comp = r.get("total_compression_pt", "?")
        print(f"  cw={cw:>6.1f} slack={slack:>+5.1f}  n_line1={n}  total_comp={comp}")
    return {
        "font_size_pt": font_size_pt,
        "n_yak": n_yak,
        "probe": probe,
        "natural": natural,
        "expected_cap": expected_cap,
        "sweep": results,
    }


def main():
    out = {}

    # Suite A: 4 font sizes × N=3
    print("\n========== SUITE A: font size × N=3 ==========")
    for fs in [10.5, 11.0, 12.0, 14.0]:
        key = f"A_fs{fs}_N3"
        out[key] = sweep_for_drop(fs, 3, "A")

    # Suite B: 12pt × N=2 mid-line (resolves Round 5 anomaly)
    print("\n========== SUITE B: N=2 mid-line at 12pt ==========")
    out["B_fs12_N2_midline"] = sweep_for_drop(12.0, 2, "B")

    os.makedirs(os.path.dirname(RESULT), exist_ok=True)
    with open(RESULT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== SUMMARY ==========")
    for key in out:
        info = out[key]
        fs = info["font_size_pt"]
        n_yak = info["n_yak"]
        natural = info["natural"]
        expected_cap = info["expected_cap"]
        sweep = info["sweep"]
        # Find max comp at PROBE_LEN chars
        max_comp = 0
        max_slack_at_24 = -1
        first_drop_slack = None
        for r in sweep:
            if r.get("n_chars_line1") == PROBE_LEN:
                tc = r.get("total_compression_pt", 0)
                if tc > max_comp:
                    max_comp = tc
                if r.get("slack", -999) > max_slack_at_24:
                    max_slack_at_24 = r["slack"]
            elif first_drop_slack is None and r.get("n_chars_line1", PROBE_LEN) < PROBE_LEN:
                if r.get("slack", -999) > 0:
                    first_drop_slack = r["slack"]
        per_yak = max_comp / n_yak if n_yak > 0 else 0
        print(f"{key}: fs={fs}pt N={n_yak}  expected_cap={expected_cap}pt")
        print(f"    max_comp_at_24={max_comp:.2f}pt  per-yak={per_yak:.3f}pt  first_drop_slack={first_drop_slack}")


if __name__ == "__main__":
    main()
