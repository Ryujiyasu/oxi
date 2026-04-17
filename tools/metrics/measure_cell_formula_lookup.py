"""Derive complete lookup table: first_para_h(fs) and extra_h(fs) for MS Mincho.

Sweeps fs 7 to 16pt. Measures row_h for n=1 (first_para_h) and n=2 (first+extra).
extra_h = row_h(n=2) - row_h(n=1).

Also compares across grids: LM0 (none), LM1 (lines), LM2 (linesAndChars).
"""
import io, json, os, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).with_name("output") / "cell_formula_lookup.json"
OUT.parent.mkdir(parents=True, exist_ok=True)
TMP = Path("pipeline_data") / "_cell_formula_tmp.docx"

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def para_xml(fs_half, text, font="ＭＳ 明朝"):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{fs_half}"/><w:szCs w:val="{fs_half}"/></w:rPr>'
    ppr = f'<w:pPr>{rpr}</w:pPr>'
    content = f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>' if text else ''
    return f'<w:p>{ppr}{content}</w:p>'


def build(paras_xml, docgrid_xml):
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


def measure_row_h(word, n_paras, fs_half, docgrid_xml):
    try: os.remove(TMP)
    except FileNotFoundError: pass
    paras_xml = [para_xml(fs_half, "AB") for _ in range(n_paras)]
    build(paras_xml, docgrid_xml)
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
            time.sleep(0.3)
            tbl = doc.Tables(1)
            tbl_top_y = tbl.Range.Information(6)
            after_rng = doc.Range(tbl.Range.End, tbl.Range.End)
            after_y = after_rng.Information(6)
            row_h = round(after_y - tbl_top_y, 2)
            doc.Close(False)
            return row_h
        except Exception:
            time.sleep(0.5 + attempt * 0.3)
    return None


def main():
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False

    grids = [
        ("none", ""),
        ("lm1_350", '<w:docGrid w:type="lines" w:linePitch="350"/>'),
        ("lm2_350", '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/>'),
    ]
    fs_list = [7.0, 8.0, 9.0, 9.5, 10.0, 10.5, 11.0, 11.5, 12.0, 13.0, 14.0, 15.0, 16.0]

    results = {}  # (grid, fs) -> {"first": row_h(n=1), "extra": row_h(n=2)-row_h(n=1)}
    try:
        for grid_label, grid_xml in grids:
            print(f"\n=== grid={grid_label} ===")
            print(f"{'fs':>5} {'row_h(1)':>10} {'row_h(2)':>10} {'first':>8} {'extra':>8}")
            for fs in fs_list:
                h1 = measure_row_h(word, 1, int(round(fs * 2)), grid_xml)
                h2 = measure_row_h(word, 2, int(round(fs * 2)), grid_xml)
                if h1 is None or h2 is None:
                    print(f"{fs:>5.1f} ERR")
                    continue
                first = h1
                extra = round(h2 - h1, 2)
                results[f"{grid_label}/{fs}"] = {"first": first, "extra": extra, "h2": h2}
                print(f"{fs:>5.1f} {h1:>10.2f} {h2:>10.2f} {first:>8.2f} {extra:>8.2f}")
    finally:
        try: word.Quit()
        except Exception: pass
        try: os.remove(TMP)
        except Exception: pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
