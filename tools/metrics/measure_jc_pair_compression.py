"""Measure Word's yakumono pair compression gated by jc (alignment)?

Hypothesis: Word applies pair half-width ONLY when para.jc == "both"
(justified). For jc=left, yakumono pairs are NOT compressed.

Minimal repro: 2-paragraph doc with identical content, different jc.
Measure per-char x via COM Information(5) to see '）' advance in each.

Expected if hypothesis CONFIRMED:
  jc=both: '）' advance ≈ 5.25pt (compressed to half)
  jc=left: '）' advance ≈ 10.5pt (full width)
"""
import io, json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP_DOCX = Path("pipeline_data") / "_jc_pair_repro.docx"
OUT = Path(__file__).with_name("output") / "jc_pair_matrix.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'

# Settings with compressPunctuation + compatMode=15
SETTINGS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:characterSpacingControl w:val="compressPunctuation"/>'
    '<w:compat>'
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    '</w:compat>'
    '</w:settings>'
)


def build(test_chars: str, include_both_jc: bool):
    """Build a doc with 2 paragraphs: jc=both and jc=left, same content."""
    rpr = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'

    def para(jc_val):
        return (
            f'<w:p><w:pPr>'
            f'<w:jc w:val="{jc_val}"/>'
            f'{rpr}</w:pPr>'
            f'<w:r>{rpr}<w:t xml:space="preserve">{test_chars}</w:t></w:r>'
            f'</w:p>'
        )

    body_paras = [para('both'), para('left')]
    body = ''.join(body_paras)
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'

    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body}{sect}</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP_DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/settings.xml", SETTINGS_XML)
        z.writestr("word/document.xml", xml)


def measure_each_paragraph(word, test_chars: str):
    """Returns list of per-paragraph per-char data.
    Each para: {"jc": "both"|"left", "chars": [{"ch","x","advance"}...]}
    """
    build(test_chars, include_both_jc=True)
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(TMP_DOCX.resolve()), ReadOnly=True)
            time.sleep(0.5)
            results = []
            jc_labels = ["both", "left"]
            for i, jc in enumerate(jc_labels):
                p = doc.Paragraphs(i + 1)
                rng = p.Range
                sel = word.Selection
                char_data = []
                for ci in range(rng.Start, rng.End):
                    sel.SetRange(ci, ci + 1)
                    try:
                        x = sel.Information(5)
                        y = sel.Information(6)
                        ch = sel.Text
                    except Exception:
                        continue
                    char_data.append({"ch": ch, "x": round(x, 2), "y": round(y, 2)})
                # Compute advance
                for j in range(len(char_data) - 1):
                    char_data[j]["adv"] = round(char_data[j+1]["x"] - char_data[j]["x"], 2)
                results.append({"jc": jc, "chars": char_data})
            doc.Close(False)
            return results
        except Exception as e:
            print(f"retry: {e}")
            time.sleep(1.0)
    return None


def main():
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False

    # Test cases: pair patterns
    test_cases = [
        ("）、ABC",   ")・・A pair"),
        ("、「ABC",   "、+「 pair"),
        ("。」ABC",   "。+」 pair"),
        ("」「ABC",   "」+「 pair"),
        ("）、）、A",  "）、 x2"),
        ("AB）、CD",  "pair in middle"),
    ]

    all_results = []
    try:
        for chars, label in test_cases:
            print(f"\n=== {label!r}: chars={chars!r} ===")
            r = measure_each_paragraph(word, chars)
            if r is None:
                print("  ERR: COM failed")
                continue
            for para in r:
                jc = para["jc"]
                # Show pair char advances
                print(f"  jc={jc}:")
                for c in para["chars"]:
                    adv = c.get("adv", "(last)")
                    print(f"    x={c['x']:>7.2f} ch={c['ch']!r} adv={adv}")
            all_results.append({"label": label, "chars": chars, "measurements": r})
    finally:
        try: word.Quit()
        except: pass
        try: os.remove(TMP_DOCX)
        except: pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved → {OUT}")

    # Summary: identify pair char advances in both vs left
    print("\n=== SUMMARY: pair char advance jc=both vs jc=left ===")
    for entry in all_results:
        print(f"\n{entry['label']} ({entry['chars']!r}):")
        m = entry["measurements"]
        # Just compare first yakumono pair position in each
        both = m[0]["chars"] if m else []
        left = m[1]["chars"] if m else []
        for j in range(min(len(both), len(left))):
            if both[j].get("adv") is not None and left[j].get("adv") is not None:
                ch = both[j]["ch"]
                adv_b = both[j]["adv"]
                adv_l = left[j]["adv"]
                diff = adv_b - adv_l
                mark = "*" if abs(diff) > 0.5 else " "
                print(f"  {mark} ch={ch!r} jc=both adv={adv_b}  jc=left adv={adv_l}  diff={diff:+.2f}")


if __name__ == "__main__":
    main()
