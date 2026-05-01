"""
§9 Footnotes — pin down the footnote line-height formula.

Existing footnote_separator.json shows lh=17.5pt for MS Mincho 10.5pt body+fn.
Spec §9.1 says "LineSpacing: 12pt (Single)" — contradicts measured 17.5pt.

This sweep varies footnote font/size and footnote lineSpacingRule:
  font/size ∈ {(MS Mincho, 10.5), (Calibri, 11), (Calibri, 10), (MS Mincho, 14)}
  fn_pPr_explicit_spacing ∈ {none, line=240 auto, line=200 exact}

Hypothesis: footnote area uses a fixed lh that is NOT derived from natural_lh.
Maybe Word's normal "Footnote Text" style applies, or there's grid-snap effect.
"""
import io, json, os, sys, time, zipfile, uuid
from pathlib import Path
import pythoncom
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).with_name("output") / "footnote_lh_sweep_phase3.json"
TMP_DIR = Path("pipeline_data") / "_footnote_lh_tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'
RELS_PKG = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>'

# Body always MS Mincho 10.5pt (constant baseline for body_area_bottom)
BODY_FONT = "ＭＳ 明朝"
BODY_SZ = 21  # 10.5pt half

# Phase 3: verify lh = max(17.5, size + 5.5) formula
FN_VARIANTS = [
    # Boundary tests: size 11/11.5/12/12.5 cross the floor
    {"label": "MSMin_11.5_default",   "fn_font": "ＭＳ 明朝", "fn_sz": 23, "fn_spacing": None},  # 11.5+5.5=17 → floor 17.5
    {"label": "MSMin_12.5_default",   "fn_font": "ＭＳ 明朝", "fn_sz": 25, "fn_spacing": None},  # 12.5+5.5=18 (above floor)
    {"label": "MSMin_15_default",     "fn_font": "ＭＳ 明朝", "fn_sz": 30, "fn_spacing": None},  # 15+5.5=20.5
    {"label": "MSMin_20_default",     "fn_font": "ＭＳ 明朝", "fn_sz": 40, "fn_spacing": None},  # 20+5.5=25.5
    {"label": "MSMin_24_default",     "fn_font": "ＭＳ 明朝", "fn_sz": 48, "fn_spacing": None},  # 24+5.5=29.5
    # Latin generalization tests
    {"label": "Calibri_14_default",   "fn_font": "Calibri",  "fn_sz": 28, "fn_spacing": None},  # 14+5.5=19.5
    {"label": "Calibri_18_default",   "fn_font": "Calibri",  "fn_sz": 36, "fn_spacing": None},  # 18+5.5=23.5
    {"label": "Calibri_12_default",   "fn_font": "Calibri",  "fn_sz": 24, "fn_spacing": None},  # 12+5.5=17.5 floor
    # Yu Mincho — has natural_lh = 1.685 × size, much larger than CJK 83/64
    {"label": "YuMincho_11_default",  "fn_font": "Yu Mincho", "fn_sz": 22, "fn_spacing": None}, # natural=18.5, vs size+5.5=16.5
    {"label": "YuMincho_14_default",  "fn_font": "Yu Mincho", "fn_sz": 28, "fn_spacing": None}, # natural=23.6, vs 19.5
]


def rpr(font, sz):
    return f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'


def build(path, variant, n_fn):
    body_runs = ""
    for i in range(n_fn):
        body_runs += (
            f'<w:r>{rpr(BODY_FONT, BODY_SZ)}<w:t xml:space="preserve">b{i+1}</w:t></w:r>'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{i+2}"/></w:r>'
        )
    body_para = f'<w:p><w:pPr>{rpr(BODY_FONT, BODY_SZ)}</w:pPr>{body_runs}</w:p>'

    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body_para}{sect}</w:body></w:document>'
    )

    fn_font, fn_sz = variant["fn_font"], variant["fn_sz"]
    spacing = ""
    if variant["fn_spacing"]:
        spacing = f'<w:spacing {variant["fn_spacing"]}/>'

    fn_ppr = f'<w:pPr>{spacing}{rpr(fn_font, fn_sz)}</w:pPr>'

    fn_entries = [
        '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for i in range(n_fn):
        fn_entries.append(
            f'<w:footnote w:id="{i+2}"><w:p>{fn_ppr}'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            f'<w:r>{rpr(fn_font, fn_sz)}<w:t xml:space="preserve"> body{i+1}</w:t></w:r></w:p></w:footnote>'
        )
    footnotes_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + "".join(fn_entries) +
        '</w:footnotes>'
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS_PKG)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/footnotes.xml", footnotes_xml)


def measure(word, path):
    last = None
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
            time.sleep(0.3)
            fns = doc.Footnotes
            ys = []
            for i in range(1, fns.Count + 1):
                ys.append(round(fns(i).Range.Information(6), 3))
            doc.Close(False)
            return ys
        except Exception as e:
            last = e
            time.sleep(0.5 + attempt * 0.5)
    raise last


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    idx = 0
    try:
        for variant in FN_VARIANTS:
            for n_fn in [1, 3, 5]:
                idx += 1
                path = TMP_DIR / f"fnlh_{idx:04d}_{uuid.uuid4().hex[:8]}.docx"
                rec = {"label": variant["label"], "fn_font": variant["fn_font"],
                       "fn_sz_half": variant["fn_sz"], "fn_size_pt": variant["fn_sz"]/2,
                       "fn_spacing": variant["fn_spacing"], "n_fn": n_fn}
                try:
                    build(path, variant, n_fn)
                    ys = measure(word, path)
                    rec["fn_ys"] = ys
                    if len(ys) >= 2:
                        rec["lh_implied"] = round(ys[1] - ys[0], 3)
                    print(f"[{idx:2d}] {variant['label']:>22} n={n_fn} -> ys={ys}, lh={rec.get('lh_implied','-')}")
                except Exception as e:
                    rec["error"] = str(e)
                    print(f"[{idx:2d}] {variant['label']}: ERR {e}")
                try:
                    path.unlink()
                except Exception:
                    pass
                results.append(rec)
    finally:
        try: word.Quit()
        except: pass
        for f in TMP_DIR.glob("*.docx"):
            try: f.unlink()
            except: pass

    OUT.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
