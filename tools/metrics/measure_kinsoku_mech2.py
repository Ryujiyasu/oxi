"""§4.6.1 round 17 — Kinsoku retreat × Mech 2 interaction.

§4.6.1 says: Type B yakumono at line-start position causes retreat
(pulled back to end of previous line) per JIS X 4051 line-start
prohibition (行頭禁則).

Question: when retreat causes line-1 to overflow content_w > cap,
what does Word do?
  (a) Honor retreat, visible overflow on line 1
  (b) Drop trigger to avoid retreat (move pre-yak char to line 2)
  (c) Multi-step retreat (more chars retreat)

Probe: 50-char line with 」 strategically placed.
  Pos 24 = 」 (Type B). Sweep cw to vary wrap point.

For cw making line 1 ≈ 23-24 chars:
  - cw=300pt (25 chars natural): no overflow, 」 at end of line 1
  - cw=288pt (24 chars natural): 」 at pos 24 fits line 1
  - cw=276pt (23 chars natural): 」 at pos 24 would start line 2 → retreat?
  - cw=264pt (22 chars natural): wrap further

Measure: where does 」 end up? Does line 1 overflow content_w?
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\kinsoku_mech2_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\kinsoku_mech2.json")
os.makedirs(OUT_DIR, exist_ok=True)

YAKUMONO_B = set("」")
PROBE_LEN = 50
FONT = "ＭＳ 明朝"


def make_probe():
    """50-char probe: 23 漢 + 」 (pos 24) + 26 漢."""
    chars = ["漢"] * PROBE_LEN
    chars[23] = "」"  # pos 24 (1-indexed)
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
            '<w:sz w:val="24"/></w:rPr>'
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
        # Group by lines
        lines_b = {}
        for t, x, y, sz in xs:
            ykey = round(y, 0)
            lines_b.setdefault(ykey, []).append((t, x, y, sz))
        line_data = []
        yak_position = None
        for ykey in sorted(lines_b.keys()):
            items = sorted(lines_b[ykey], key=lambda v: v[1])
            n = len(items)
            chars_text = "".join(it[0] for it in items)
            advs = []
            for i in range(len(items) - 1):
                advs.append((items[i][0], round(items[i+1][1] - items[i][1], 3),
                             items[i][3]))
            # Find 」
            yak_in_line = None
            for i, (ch, adv, sz) in enumerate(advs):
                if ch == "」":
                    yak_in_line = {"line_pos": i+1, "adv": adv, "sz": sz}
                    break
            line_width = (items[-1][1] - items[0][1]) + items[-1][3] if items else 0
            line_data.append({
                "y": ykey,
                "n_chars": n,
                "text_summary": chars_text[:50],
                "line_width": round(line_width, 2),
                "yak_in_line": yak_in_line,
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
    print(f"probe: {probe!r}")
    print(f"  pos 24 = 」 (Type B)")
    print(f"  natural = {natural}pt")

    # cw values to test wrap behavior around 」 position
    # 24 chars natural = 288pt, 25 chars = 300pt
    cw_values = [
        300,    # exact 25 chars/line, 」 at line 1 end
        296,    # slack 4 (within Mech 2 range)
        294,    # slack 6 (cap)
        292,    # slack 8 (over cap, drop expected)
        288,    # exactly 24 chars, would 」 retreat?
        284,    # 1 char drop, 23 chars line 1
        276,    # 23 chars natural (12*23=276)
        264,    # 22 chars natural
    ]

    for cw in cw_values:
        label = f"K_cw{cw}"
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
        print(f"\n[cw={cw}]")
        for li, ln in enumerate(r.get("lines", []), start=1):
            yak_str = ""
            if ln.get("yak_in_line"):
                ya = ln["yak_in_line"]
                yak_str = f"  [」 at L{li} pos {ya['line_pos']} adv={ya['adv']:.2f}]"
            print(f"  L{li}: n_chars={ln['n_chars']:>3} width={ln['line_width']:>6.2f}  vs cw={cw}  diff={ln['line_width']-cw:+5.2f}{yak_str}")

        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
