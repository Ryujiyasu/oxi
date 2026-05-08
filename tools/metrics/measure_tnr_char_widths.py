"""Measure Word's actual Times New Roman 10.5pt char widths via per-char COM x position.

Generate a 1-paragraph docx with chars laid out: " AaBb1." (each separated by
nothing, just consecutive). Use Word COM Information(WD_HPOS) on each Character
to compute char width = next.x - this.x.

Output JSON for analysis vs com_tw_overrides values.
"""
from __future__ import annotations

import json
import os
import sys
import zipfile

import win32com.client

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.join(REPO, "tools", "golden-test", "repros", "grid_snap")

WD_HPOS = 5
WD_VPOS = 6


def make_repro_docx(label: str, sample_text: str) -> str:
    """Create a minimal docx with one paragraph of sample_text in TNR 10.5pt."""
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="21"/></w:rPr>
<w:t xml:space="preserve">{sample_text}</w:t></w:r>
</w:p>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>
<w:docGrid w:linePitch="0"/>
</w:sectPr>
</w:body>
</w:document>"""

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

    settings = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/settings.xml", settings)
        zf.writestr("word/styles.xml", styles)
        zf.writestr("word/document.xml", document_xml)
    return out_path


def measure_widths(word, docx_path: str, expected_text: str) -> list[dict]:
    doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False)
    results = []
    try:
        chars = doc.Paragraphs(1).Range.Characters
        n = chars.Count
        for ci in range(1, n + 1):
            c = chars(ci)
            ch = c.Text
            if ch == '\r' or ch == '\x07':
                continue
            try:
                x = c.Information(WD_HPOS)
                y = c.Information(WD_VPOS)
                results.append({
                    "ci": ci,
                    "ch": ch,
                    "cp": ord(ch) if len(ch) == 1 else -1,
                    "x": round(x, 3),
                    "y": round(y, 3),
                })
            except Exception as e:
                results.append({"ci": ci, "ch": ch, "error": str(e)})
    finally:
        doc.Close(SaveChanges=False)

    # Compute width per char = next.x - this.x
    for i in range(len(results) - 1):
        if 'x' in results[i] and 'x' in results[i+1]:
            # Only valid if same y (= same line)
            if abs(results[i]['y'] - results[i+1]['y']) < 0.5:
                results[i]['width'] = round(results[i+1]['x'] - results[i]['x'], 3)
                results[i]['width_tw'] = round((results[i+1]['x'] - results[i]['x']) * 20, 1)
    return results


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    # Pad each char with another so we can subtract x positions
    # Use a sequence with each interesting char twice (XX) to measure width
    # Simpler: use long ASCII string and measure consecutive x diffs
    sample = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z a b c d e f g h i j k l m n o p q r s t u v w x y z 0 1 2 3 4 5 6 7 8 9 . , ; : ! ? / ( ) "

    docx = make_repro_docx("tnr_widths_sample", sample)
    print(f"Generated: {docx}")

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        results = measure_widths(word, docx, sample)
    finally:
        word.Quit()

    # Per-cp width aggregation
    by_cp = {}
    for r in results:
        if 'cp' in r and 'width' in r:
            cp = r['cp']
            if cp not in by_cp:
                by_cp[cp] = []
            by_cp[cp].append(r['width'])

    # Compare to com_tw_overrides
    with open(os.path.join(REPO, "crates/oxidocs-core/src/font/data/com_tw_overrides.json"), encoding='utf-8') as f:
        com = json.load(f)
    tnr = com.get('Times New Roman', {}).get('10.5', {})

    with open(os.path.join(REPO, "crates/oxidocs-core/src/font/data/font_metrics_compact.json"), encoding='utf-8') as f:
        fm_list = json.load(f)
    tnr_metrics = next((m for m in fm_list if m['family'] == 'Times New Roman'), None)
    upm = tnr_metrics['units_per_em']
    widths_em = {int(cp): adv/upm for cp, adv in tnr_metrics.get('widths', {}).items()}

    print(f"\n{'cp':>4} {'ch':>3} {'word_w':>8} {'com_tw':>8} {'formula':>8}")
    print('-' * 40)
    for cp in sorted(by_cp.keys()):
        ws = by_cp[cp]
        word_w_pt = sum(ws) / len(ws)
        word_w_tw = round(word_w_pt * 20, 1)
        com_v = tnr.get(str(cp), '-')
        adv_em = widths_em.get(cp, '-')
        if isinstance(adv_em, float):
            formula = int(adv_em * 10.5 * 2 + 0.5) * 10
        else:
            formula = '-'
        ch = chr(cp) if 32 <= cp < 127 else f"U+{cp:04X}"
        print(f"{cp:>4} {ch!r:>4} {word_w_tw:>8.1f} {str(com_v):>8s} {str(formula):>8s}")

    # Save raw measurements
    OUT = os.path.join(REPO, "pipeline_data", "tnr_word_actual_widths.json")
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"font": "Times New Roman", "size": 10.5,
                    "measurements": results,
                    "by_cp": {str(k): v for k, v in by_cp.items()}},
                   f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {OUT}")


if __name__ == "__main__":
    main()
