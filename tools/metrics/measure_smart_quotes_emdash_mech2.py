"""§4.7 round 11 — smart quotes + em-dash + hbar Mech 2 fire verification.

§4.7 lists:
  Type A: ‘ (U+2018), " (U+201C)
  Type B: ’ (U+2019), " (U+201D), — (U+2014 em-dash)
  Type C: ― (U+2015 hbar)

Session 51 found Mech 1 treats em-dash as Type B for MS branded fonts
but Type C for Yu Mincho, Meiryo. Hbar universal Type C.

Round 11: verify Mech 2 (cSC=compressPunctuation slack distribution)
behavior for these chars across fonts.

Test:
  Suite A: 6 chars × MS Mincho 12pt × 4 slack values
    Chars: ‘ ’ " " — ―
    Slacks: -1 (no overflow), +2, +4, +6 (cap)
  Suite B: em-dash × 3 fonts × 1 overflow cw
    Fonts: MS Mincho, Yu Mincho, Meiryo
    Detect: is Mech 2 font-dependent for em-dash like Mech 1?
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\quotes_emdash_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\quotes_emdash_mech2.json")
os.makedirs(OUT_DIR, exist_ok=True)

# Compressible chars per spec §4.7 (we'll test each individually)
TEST_CHARS = {
    "LSQ_2018":  "‘",   # ‘  Type A
    "RSQ_2019":  "’",   # ’  Type B
    "LDQ_201C":  "“",   # "  Type A
    "RDQ_201D":  "”",   # "  Type B
    "EmDash_2014": "—", # —  Type B (font-dependent per Session 51)
    "Hbar_2015": "―",   # ―  Type C (universal no compress)
    # Reference: standard yakumono
    "RBracket_300D": "」",  # 」  Type B (control)
}

PROBE_LEN = 24


def make_probe(test_char):
    """24-char probe with test_char at pos 12."""
    chars = ["漢"] * PROBE_LEN
    chars[11] = test_char
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


def measure_one(path, target_char):
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
        # Find target char advance
        target_adv = None
        target_ratio = None
        target_sz = None
        for i in range(len(line1) - 1):
            t, _, sz = line1[i]
            if t == target_char:
                adv = round(line1[i+1][1] - line1[i][1], 3)
                target_adv = adv
                target_sz = sz
                target_ratio = round(adv / sz, 4) if sz > 0 else None
                break
        # Total line compression (any char shorter than its size)
        total_comp = sum((sz - (line1[i+1][1] - line1[i][1]))
                         for i, (t, _, sz) in enumerate(line1[:-1])
                         if sz > 0 and (line1[i+1][1] - line1[i][1]) < sz * 0.99)
        return {
            "n_chars_line1": n_line1,
            "target_char_advance": target_adv,
            "target_char_size": target_sz,
            "target_ratio": target_ratio,
            "is_compressed": (target_ratio is not None and target_ratio < 0.99),
            "total_line_compression": round(total_comp, 3),
        }
    finally:
        try: word.Quit()
        except: pass


SUITE_A_FONT = "ＭＳ 明朝"
SUITE_A_SZ = 24  # 12pt
SUITE_A_NATURAL = PROBE_LEN * 12.0   # 288pt
SUITE_A_SLACKS = [-1.0, 2.0, 4.0, 6.0]   # no overflow, mild, partial, cap

SUITE_B_FONTS = ["ＭＳ 明朝", "Yu Mincho", "Meiryo"]
SUITE_B_CHAR = "—"   # em-dash
SUITE_B_LABEL = "EmDash_2014"
SUITE_B_SLACK = 4.0


def main():
    out = {}

    # Suite A: 7 chars × 4 slack values at MS Mincho 12pt
    print("\n========== Suite A: char × slack at MS Mincho 12pt ==========")
    for char_label, ch in TEST_CHARS.items():
        probe = make_probe(ch)
        for slack in SUITE_A_SLACKS:
            cw = round(SUITE_A_NATURAL - slack, 1)
            label = f"A_{char_label}_slack{slack:.0f}"
            try:
                p = make_docx(label, probe, cw, SUITE_A_FONT, SUITE_A_SZ, "both")
            except Exception as e:
                out[label] = {"build_error": str(e)}
                continue
            kill_word()
            try:
                r = measure_one(p, ch)
            except Exception as e:
                r = {"measure_error": str(e)}
                kill_word()
            out[label] = {"char": char_label, "char_unicode": f"U+{ord(ch):04X}",
                          "slack": slack, "font": SUITE_A_FONT, **r}
            adv = r.get("target_char_advance")
            ratio = r.get("target_ratio")
            n = r.get("n_chars_line1", "?")
            comp = "comp" if r.get("is_compressed") else "no  "
            adv_str = f"{adv:.2f}" if adv is not None else "N/A"
            ratio_str = f"{ratio:.3f}" if ratio is not None else "N/A"
            print(f"  [{char_label:<14}] slack={slack:>+5.1f} n={n} target adv={adv_str} ratio={ratio_str} {comp}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)

    # Suite B: em-dash × 3 fonts at slack=4
    print("\n========== Suite B: em-dash × 3 fonts at slack=4 ==========")
    probe_em = make_probe(SUITE_B_CHAR)
    for font in SUITE_B_FONTS:
        cw = round(SUITE_A_NATURAL - SUITE_B_SLACK, 1)
        font_id = font.replace(" ", "").replace("ＭＳ", "MS").replace("明朝", "Min")
        label = f"B_{SUITE_B_LABEL}_{font_id}_slack4"
        try:
            p = make_docx(label, probe_em, cw, font, SUITE_A_SZ, "both")
        except Exception as e:
            out[label] = {"build_error": str(e)}
            continue
        kill_word()
        try:
            r = measure_one(p, SUITE_B_CHAR)
        except Exception as e:
            r = {"measure_error": str(e)}
            kill_word()
        out[label] = {"char": SUITE_B_LABEL, "char_unicode": f"U+{ord(SUITE_B_CHAR):04X}",
                      "slack": SUITE_B_SLACK, "font": font, **r}
        adv = r.get("target_char_advance")
        ratio = r.get("target_ratio")
        comp = "FIRES" if r.get("is_compressed") else "no   "
        adv_str = f"{adv:.2f}" if adv is not None else "N/A"
        ratio_str = f"{ratio:.3f}" if ratio is not None else "N/A"
        print(f"  [{font:<15}] adv={adv_str} ratio={ratio_str} {comp}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    # Summary
    print("\n========== SUMMARY: Type A/B/C verdict per char (Mech 2) ==========")
    print(f"{'Char (unicode)':<20} {'expected':<10} {'fires Mech 2?':<15} {'final type':<12}")
    for char_label, ch in TEST_CHARS.items():
        # At slack=4 with MS Mincho, did Mech 2 fire?
        key = f"A_{char_label}_slack4"
        info = out.get(key, {})
        compressed = info.get("is_compressed", False)
        unicode = f"U+{ord(ch):04X}"
        expected = {
            "LSQ_2018": "A (compresses)",
            "RSQ_2019": "B (compresses)",
            "LDQ_201C": "A (compresses)",
            "RDQ_201D": "B (compresses)",
            "EmDash_2014": "B (compresses)",
            "Hbar_2015": "C (no compress)",
            "RBracket_300D": "B (compresses)",
        }.get(char_label, "?")
        verdict = "FIRES" if compressed else "no"
        final = "A or B" if compressed else "C"
        print(f"  {char_label} ({unicode}):<{ch}>  {expected:<25} {verdict:<6}  {final}")


if __name__ == "__main__":
    main()
