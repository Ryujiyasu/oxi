"""§4.6.2 round 27 — cross-font RSB validation.

Round 25/26 derived RSB formula at fs=12 Times New Roman:
  boundary_advance = nat_advance + 3.0pt - rsb(glyph)

Round 27 verifies the formula scope on other Latin fonts (Arial, Calibri)
at fs=12pt. If the formula is truly RSB-based, different fonts should
show different RSB values (consistent with the font's design metrics).

Test glyphs: a, M, 1, n, T (5 glyphs from Round 25).
Latin fonts: Arial (sans), Calibri (sans modern).
CJK font: ＭＳ 明朝 (constant).
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\autospace_cross_font_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\autospace_cross_font.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(probe, sz_val, page_w_tw, margin_tw, latin_font):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{latin_font}" w:hAnsi="{latin_font}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t xml:space="preserve">{probe}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml(sz_val, latin_font):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{latin_font}" w:hAnsi="{latin_font}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '</w:settings>')


def make_docx(label, probe, sz_val, latin_font):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((400 + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(probe, sz_val, page_w_tw, margin_tw, latin_font)
    styles_xml = make_styles_xml(sz_val, latin_font)
    settings_xml = make_settings_xml()
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
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
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
        ' Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
        ' Target="settings.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", styles_xml)
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
                    xs.append((t, float(c.Information(5))))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        char_advances = []
        for i in range(len(xs) - 1):
            t = xs[i][0]
            adv = round(xs[i+1][1] - xs[i][1], 3)
            char_advances.append({"ch": t, "adv": adv})
        return {"n_chars": len(xs), "char_advances": char_advances}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    glyphs = ["a", "M", "1", "n", "T"]
    fs_pt = 12.0
    sz_val = 24
    extra_pt = 3.0  # §4.6.2 at fs=12

    fonts = ["Arial", "Calibri"]

    for latin_font in fonts:
        print(f"\n=== Latin font: {latin_font} fs={fs_pt}pt extra={extra_pt}pt ===")
        out[latin_font] = {"font": latin_font, "fs": fs_pt, "tests": []}
        for ch in glyphs:
            natural_adv = None
            boundary_adv = None
            for probe_kind, probe in [("nat", ch + ch), ("bnd", ch + "漢")]:
                label = f"{latin_font}_{ch}_{probe_kind}"
                try:
                    p = make_docx(label, probe, sz_val, latin_font)
                except Exception as e:
                    continue
                kill_word()
                try:
                    r = measure_one(p)
                except Exception as e:
                    r = {"measure_error": str(e)}
                    kill_word()
                advs = r.get("char_advances", [])
                if advs:
                    if probe_kind == "nat":
                        natural_adv = advs[0]["adv"]
                    else:
                        boundary_adv = advs[0]["adv"]
            extra = None
            implied_rsb = None
            if natural_adv is not None and boundary_adv is not None:
                extra = round(boundary_adv - natural_adv, 3)
                implied_rsb = round(extra_pt - extra, 3)
            entry = {
                "ch": ch,
                "natural_adv": natural_adv,
                "boundary_adv": boundary_adv,
                "extra": extra,
                "implied_rsb_pt": implied_rsb,
            }
            out[latin_font]["tests"].append(entry)
            print(f"  {ch!r}: nat={natural_adv} bnd={boundary_adv} extra={extra} → rsb={implied_rsb}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== SUMMARY ==========")
    print(f"{'font':>10} {'glyph':>6} {'nat':>6} {'bnd':>6} {'extra':>7} {'implied_rsb':>12}")
    for fnt in fonts:
        info = out.get(fnt, {})
        for t in info.get("tests", []):
            print(f"  {fnt:>9} {t['ch']!r:>5} "
                  f"{t.get('natural_adv','?'):>6} {t.get('boundary_adv','?'):>6} "
                  f"{t.get('extra','?'):>7} {t.get('implied_rsb_pt','?'):>12}")


if __name__ == "__main__":
    main()
