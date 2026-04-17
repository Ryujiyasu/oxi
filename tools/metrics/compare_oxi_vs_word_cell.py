"""Compare Oxi vs Word row heights on controlled minimal-repro cells.

For each test case:
1. Build a docx with single-cell table (tcMar=0, no borders)
2. Measure Word row_h via COM
3. Render with Oxi (--dump-layout), compute row_h from border positions
4. Report Δ = Oxi - Word to localize the per-cell gap

Tests single-fs cells at various n_paras and font sizes.
"""
import io, json, os, subprocess, sys, time, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OXI_RENDERER = Path(r"C:/Users/ryuji/oxi-4/tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
TMP_DOCX = Path("pipeline_data") / "_compare_tmp.docx"
TMP_JSON = Path("pipeline_data") / "_compare_tmp.json"

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def para_xml(fs_half, text, font="ＭＳ 明朝"):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{fs_half}"/><w:szCs w:val="{fs_half}"/></w:rPr>'
    ppr = f'<w:pPr>{rpr}</w:pPr>'
    content = f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>' if text else ''
    return f'<w:p>{ppr}{content}</w:p>'


def build(paras_spec, docgrid_xml):
    paras_xml = [para_xml(hp, txt) for (hp, txt) in paras_spec]
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
    with zipfile.ZipFile(TMP_DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def measure_word_row_h(word, paras_spec, docgrid_xml):
    """Open the currently-built TMP_DOCX and return row_h via COM."""
    try:
        doc = word.Documents.Open(str(TMP_DOCX.resolve()), ReadOnly=True)
        time.sleep(0.4)
        tbl = doc.Tables(1)
        tbl_top_y = tbl.Range.Information(6)
        after_rng = doc.Range(tbl.Range.End, tbl.Range.End)
        after_y = after_rng.Information(6)
        row_h = after_y - tbl_top_y
        doc.Close(False)
        return round(row_h, 2)
    except Exception as e:
        print(f"    word err: {e}")
        return None


def measure_oxi_row_h():
    """Render TMP_DOCX with Oxi and extract row_h from border elements."""
    try: os.remove(TMP_JSON)
    except FileNotFoundError: pass
    trash = str(Path("pipeline_data") / "_compare_trash").replace("\\", "/")
    result = subprocess.run([str(OXI_RENDERER), str(TMP_DOCX.resolve()), trash, "150", f"--dump-layout={TMP_JSON.resolve()}"],
                            capture_output=True, timeout=30)
    if result.returncode != 0:
        err = result.stderr.decode('utf-8', errors='replace')[:200]
        print(f"    oxi err: {err}")
        return None
    if not TMP_JSON.exists():
        return None
    d = json.load(open(TMP_JSON, encoding='utf-8'))
    p1 = d['pages'][0]['elements']
    h_bords = [e for e in p1 if e.get('type') == 'border' and e.get('h', 0) == 0.0 and e.get('w', 0) > 0]
    ys = sorted(set(b['y'] for b in h_bords))
    if len(ys) >= 2:
        return round(ys[1] - ys[0], 2)
    return None


def main():
    grids = [
        ("none", ""),
        ("lm2_350", '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/>'),
    ]

    test_cases = []
    for fs_hp in [18, 21, 24]:  # 9, 10.5, 12 pt
        for n in [1, 2, 3]:
            paras = [(fs_hp, f"P{i+1}") for i in range(n)]
            test_cases.append((f"n{n}_fs{fs_hp/2}", paras))

    # Build all docx files first, then measure Word, then measure Oxi.
    # Word COM single-instance is more stable this way.
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False

    all_rows = []  # (label, grid, word_h, oxi_h, delta)
    try:
        for grid_label, grid_xml in grids:
            print(f"\n=== grid={grid_label} ===")
            print(f"{'case':<15} {'word':>8} {'oxi':>8} {'Δ':>8}")
            for label, paras in test_cases:
                build(paras, grid_xml)
                wh = measure_word_row_h(word, paras, grid_xml)
                if wh is None:
                    print(f"{label:<15} WORD_ERR")
                    all_rows.append({"label": label, "grid": grid_label, "word": None, "oxi": None})
                    continue
                oh = measure_oxi_row_h()
                if oh is None:
                    print(f"{label:<15} OXI_ERR (word={wh})")
                    all_rows.append({"label": label, "grid": grid_label, "word": wh, "oxi": None})
                    continue
                delta = oh - wh
                mark = "!" if abs(delta) > 0.3 else " "
                print(f"{mark} {label:<15} {wh:>8.2f} {oh:>8.2f} {delta:>+8.2f}")
                all_rows.append({"label": label, "grid": grid_label, "word": wh, "oxi": oh, "delta": round(delta, 2)})
    finally:
        try: word.Quit()
        except Exception: pass
        try: os.remove(TMP_DOCX)
        except Exception: pass
        try: os.remove(TMP_JSON)
        except Exception: pass

    OUT = Path(__file__).with_name("output") / "compare_oxi_vs_word_cell.json"
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(all_rows, f, indent=2)
    print(f"\nSaved → {OUT}")


if __name__ == "__main__":
    main()
