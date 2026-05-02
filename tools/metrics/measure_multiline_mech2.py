"""§4.7b round 9 — multi-line Mech 2 cascade.

Question: when paragraph wraps to multiple lines (force wrap), does
Mech 2 cap apply:
  (i)  per-line independently (each line has own cap = floor(sz/2)*0.5)
  (ii) cumulatively across lines (paragraph-level budget)
  (iii) some other interaction (e.g., re-distribute slack across lines)

Test design:
  Probe: 50-char line with 6 yakumono distributed evenly.
    yakumono at positions 5, 12, 20, 30, 38, 45 (covering both halves).
  cw values: tight (forces wrap to 2 lines, with overflow on each).

For each cw, measure per-line:
  - n_chars
  - total_compression
  - which yak compressed by how much

Compare to single-line cap formula prediction.

Probe at MS Mincho 12pt:
  natural = 50 × 12 = 600pt
  Test cw values: 264 (~22 chars/line), 280, 300, 320 (varying line counts)
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\multiline_mech2_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\multiline_mech2.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO = set("「")
PROBE_LEN = 50
FONT = "ＭＳ 明朝"
SIZE_HALF = 24  # 12pt


def make_probe():
    """50-char probe, 6 yakumono evenly distributed."""
    chars = ["漢"] * PROBE_LEN
    for p in [4, 11, 19, 29, 37, 44]:  # 0-indexed → 1-indexed: 5,12,20,30,38,45
        chars[p] = "「"
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
            f'<w:sz w:val="{SIZE_HALF}"/></w:rPr>'
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
                    if t in ("\r", "\x07"): continue
                    xs.append((t, float(c.Information(5)), float(c.Information(6)),
                               float(c.Font.Size if c.Font.Size else 0)))
                except Exception: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        # Group by lines (y-bucketed)
        lines_b = {}
        for t, x, y, sz in xs:
            ykey = round(y, 0)
            lines_b.setdefault(ykey, []).append((t, x, y, sz))
        line_data = []
        for ykey in sorted(lines_b.keys()):
            items = sorted(lines_b[ykey], key=lambda v: v[1])
            n_line = len(items)
            advs = []
            for i in range(len(items) - 1):
                advs.append((items[i][0], round(items[i+1][1] - items[i][1], 3),
                             items[i][3]))
            n_yak = sum(1 for t, _, sz in advs if t in YAKUMONO)
            n_yak_comp = sum(1 for t, a, sz in advs if t in YAKUMONO and sz > 0 and a < sz * 0.99)
            total_comp = sum((sz - a) for t, a, sz in advs if t in YAKUMONO and sz > 0 and a < sz * 0.99)
            comp_yak_detail = [(t, round(a, 2), round(sz, 1)) for t, a, sz in advs
                               if t in YAKUMONO and sz > 0 and a < sz * 0.99]
            line_data.append({
                "y": ykey,
                "n_chars": n_line,
                "n_yak": n_yak,
                "n_yak_compressed": n_yak_comp,
                "total_compression_pt": round(total_comp, 3),
                "comp_yak": comp_yak_detail,
            })
        return {
            "n_lines": len(line_data),
            "lines": line_data,
        }
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    probe = make_probe()
    natural = PROBE_LEN * 12.0  # 600pt
    yak_positions = [i+1 for i, c in enumerate(probe) if c in YAKUMONO]
    print(f"probe: {probe!r}")
    print(f"yak positions (1-indexed): {yak_positions}")
    print(f"natural = {natural}pt")

    # Test cw values forcing wrap to 2-3 lines
    cw_values = [
        # 1-line scenarios
        # 600 (natural, no overflow)
        # 590 (slack 10 for whole 50-char line)

        # 2-line wrap scenarios - force ~25 chars per line
        310, 308, 306, 304, 302, 300,  # ~25 chars/line, 2 lines
        # 3-line wrap scenarios - force ~17 chars per line
        220, 218, 216, 214, 212, 210,
    ]

    for cw in cw_values:
        label = f"ML_cw{cw}"
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
        out[label] = {"content_w": cw, "natural": natural, **r}
        n_lines = r.get("n_lines", "?")
        print(f"\n[{label}] cw={cw}pt natural={natural}pt n_lines={n_lines}")
        if "lines" in r:
            for li, ln in enumerate(r["lines"], start=1):
                print(f"  L{li}: n_chars={ln['n_chars']:>3} n_yak={ln['n_yak']} comp={ln['n_yak_compressed']}  total_comp={ln['total_compression_pt']:>5.2f}pt  detail={ln['comp_yak']}")
        # Incremental save
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    print(f"\nWrote {RESULT}")


if __name__ == "__main__":
    main()
