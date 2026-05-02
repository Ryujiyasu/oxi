"""§4.6.2 round 25 — per-glyph Latin→kana boundary extra at fs=12.

Round 24 finding: at fs=12 Times New Roman, kana→Latin boundary
extra is constant +3.0pt, but Latin→kana boundary extra varies by
glyph: M=+3.0, a=+2.0, 1=+2.5.

Hypothesis: Latin→kana extra depends on Latin glyph's natural width
(narrower glyph → smaller extra).

Round 25 sweep:
  Letters: i (narrow), l, c, e, a, n, o, M (wide), W, T, A, x, w
  Digits: 0..9
  Format per glyph G: probe = "G漢" — measure G's adv vs natural G's adv.
  natural G_adv: probe = "GG" — measure first G's adv (the 2nd G has no boundary).

For each glyph:
  natural_adv = (probe "GG" first G adv)
  boundary_adv = (probe "G漢" first G adv) — boundary applied
  extra = boundary_adv - natural_adv
  ratio = extra / 3.0 (= extra normalized to baseline kana→Latin extra)
  width_pt = natural_adv

Look for: extra(G) = f(natural_width(G))?
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\autospace_glyph_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\autospace_glyph.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"
LATIN_FONT = "Times New Roman"


def make_doc_xml(probe, sz_val, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="left"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t xml:space="preserve">{probe}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml(sz_val):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{LATIN_FONT}" w:hAnsi="{LATIN_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '</w:settings>')


def make_docx(label, probe, sz_val):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    page_w_tw = int((400 + 170) * 20)
    margin_tw = 170 * 10
    doc_xml = make_doc_xml(probe, sz_val, page_w_tw, margin_tw)
    styles_xml = make_styles_xml(sz_val)
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
    sz_val = 24

    # Test glyphs: representative narrow → wide letters + digits
    glyphs = [
        # letters narrow → wide
        "i", "l", "t", "c", "e", "a", "n", "o", "m", "w",
        # uppercase
        "A", "M", "T", "W",
        # digits
        "0", "1", "4", "8",
    ]

    for ch in glyphs:
        # Probe 1: GG to get natural advance of first G (no boundary)
        # Probe 2: G漢 to get boundary advance of G
        natural_adv = None
        boundary_adv = None

        for probe_kind, probe in [("nat", ch + ch), ("bnd", ch + "漢")]:
            label = f"g_{ord(ch)}_{probe_kind}"
            try:
                p = make_docx(label, probe, sz_val)
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
        if natural_adv is not None and boundary_adv is not None:
            extra = round(boundary_adv - natural_adv, 3)

        out[ch] = {
            "ch": ch,
            "ord": ord(ch),
            "natural_adv": natural_adv,
            "boundary_adv": boundary_adv,
            "extra": extra,
        }
        print(f"  {ch!r}: nat={natural_adv} bnd={boundary_adv} extra={extra}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== SUMMARY (sorted by natural_adv) ==========")
    print(f"{'ch':>4} {'ord':>5} {'natural':>10} {'boundary':>10} {'extra':>10}")
    valid = [(k, v) for k, v in out.items() if v.get("extra") is not None]
    valid.sort(key=lambda kv: kv[1].get("natural_adv", 0))
    for k, v in valid:
        nat = v["natural_adv"]
        bnd = v["boundary_adv"]
        ext = v["extra"]
        print(f"  {k!r:>3} {v['ord']:>5} {nat:>10.2f} {bnd:>10.2f} {ext:>10.2f}")


if __name__ == "__main__":
    main()
