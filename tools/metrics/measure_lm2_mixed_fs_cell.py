"""Mixed-fs cell row height sweep for LM2 linesAndChars_350 (b35 grid).

Tests cells with multiple paragraphs at different font sizes, including
trailing empty pilcrows. This extends measure_lm2_multipara_cell.py which
only tested single-fs stacks.

Hypothesized b35 case: cell with [content fs=10.5] + [subtitle fs=9] +
[empty fs=12] — each paragraph contributes differently to row height.
"""
import io, json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).with_name("output") / "lm2_mixed_fs_cell.json"
OUT.parent.mkdir(parents=True, exist_ok=True)
TMP = Path("pipeline_data") / "_lm2_mixed_fs_tmp.docx"

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def para_xml(fs_half, text, font="ＭＳ 明朝"):
    """Create paragraph XML with given fs (half-pts) and text."""
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{fs_half}"/><w:szCs w:val="{fs_half}"/></w:rPr>'
    ppr = f'<w:pPr>{rpr}</w:pPr>'
    content = f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>' if text else ''
    return f'<w:p>{ppr}{content}</w:p>'


def build(paras_xml):
    """Build doc with a single-cell table containing paras_xml."""
    docgrid_xml = '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/>'
    tbl = (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="4000" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        f'{"".join(paras_xml)}'
        '</w:tc></w:tr>'
        '</w:tbl>'
    )
    sentinel = para_xml(21, "END")
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>{docgrid_xml}</w:sectPr>'
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{tbl}{sentinel}{sect}</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def measure(word, paras_spec, label):
    """paras_spec = list of (fs_half_pt, text_or_empty). Returns row_h + para_ys."""
    try: os.remove(TMP)
    except FileNotFoundError: pass
    paras_xml = [para_xml(fs_hp, txt) for (fs_hp, txt) in paras_spec]
    build(paras_xml)
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
            time.sleep(0.3)
            tbl = doc.Tables(1)
            tbl_top_y = tbl.Range.Information(6)
            after_rng = doc.Range(tbl.Range.End, tbl.Range.End)
            after_y = after_rng.Information(6)
            row_h = after_y - tbl_top_y
            para_ys = []
            for i in range(1, len(paras_spec) + 1):
                p_y = doc.Paragraphs(i).Range.Information(6)
                para_ys.append(p_y)
            doc.Close(False)
            return {
                "label": label,
                "paras": [(hp/2.0, txt) for (hp,txt) in paras_spec],
                "row_h": round(row_h, 2),
                "para_ys": [round(y, 2) for y in para_ys],
                "gaps": [round(para_ys[i+1]-para_ys[i], 2) for i in range(len(para_ys)-1)],
            }
        except Exception as e:
            time.sleep(0.8 + attempt * 0.5)
    return {"label": label, "error": "com retry exhausted"}


TEST_CASES = [
    ("1x10.5 content", [(21, "AB")]),
    ("1x10.5 empty", [(21, "")]),
    ("2x10.5 content", [(21, "AB"), (21, "CD")]),
    ("1x10.5+1x12 empty", [(21, "AB"), (24, "")]),
    ("1x10.5+1x9 empty", [(21, "AB"), (18, "")]),
    ("1x10.5+1x9 content", [(21, "AB"), (18, "CD")]),
    ("1x10.5+1x12 content", [(21, "AB"), (24, "CD")]),
    ("1x10.5+1x10.5 empty", [(21, "AB"), (21, "")]),
    ("1x10.5+1x9+1x12 empty", [(21, "AB"), (18, "CD"), (24, "")]),  # b35 row 7 pattern
    ("1x10.5+1x9 content+1x12 empty", [(21, "AB"), (18, "CD"), (24, "")]),
    ("1x10.5 content+1x12 empty only", [(21, "AB"), (24, "")]),  # b35 row 6 style
    ("3x10.5 content", [(21, "AB"), (21, "CD"), (21, "EF")]),
    ("3x empty", [(21, ""), (21, ""), (21, "")]),
    ("1x10.5+1x10.5 empty+1x12 empty", [(21, "AB"), (21, ""), (24, "")]),  # b35 row 3 pattern
]


def main():
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    try:
        for label, spec in TEST_CASES:
            r = measure(word, spec, label)
            results.append(r)
            if "error" in r:
                print(f"  [ERR] {label}: {r['error']}")
            else:
                specs_str = ", ".join(f"{fs}pt{'Ø' if not txt else ''}" for (fs,txt) in r['paras'])
                print(f"  {label:<40}: row_h={r['row_h']:>6.2f}  specs=[{specs_str}]  gaps={r['gaps']}")
    finally:
        try: word.Quit()
        except Exception: pass
        try: os.remove(TMP)
        except Exception: pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
