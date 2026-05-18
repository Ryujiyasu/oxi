"""S105: per-row Oxi vs Word diff on bottom-band target docs.

For a1d6 / d4d126 / de6e, capture per-row pitch diff (Oxi vs Word).
Identify which specific rows mismatch and what XML attributes they have.
This is more targeted than full baseline survey.
"""
import json, os, re, subprocess, sys, tempfile, zipfile
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
DOCX_DIR = ROOT / 'tools/golden-test/documents/docx'
RENDERER = ROOT / 'tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe'
OUT = ROOT / 'tools/metrics/per_row_diff_bottom_band.json'

TARGETS = [
    'a1d6e4efa2e7_tokumei_08_01-4.docx',
    'd4d126dfe1d9_tokumei_08_01-3.docx',
    'de6e32b5960b_tokumei_08_01-1.docx',
]


def parse_rows_xml(docx_path: Path) -> dict:
    """Parse OOXML for per-row metadata. Return {(table_idx, row_idx): meta}."""
    with zipfile.ZipFile(docx_path) as z:
        doc = z.read('word/document.xml').decode('utf-8', errors='ignore')
    result = {}
    pos = 0
    t_idx = 0
    while True:
        ts = doc.find('<w:tbl>', pos)
        if ts < 0:
            ts = doc.find('<w:tbl ', pos)
            if ts < 0:
                break
        te = doc.find('</w:tbl>', ts)
        if te < 0:
            break
        tbl_xml = doc[ts:te]
        r_idx = 0
        rpos = 0
        while True:
            trm = re.search(r'<w:tr(?:\s|>)', tbl_xml[rpos:])
            if not trm:
                break
            rs = rpos + trm.start()
            re_pos = tbl_xml.find('</w:tr>', rs)
            if re_pos < 0:
                break
            tr_xml = tbl_xml[rs:re_pos]
            tr_meta = {
                'h_rule': None,
                'trH_val': None,
                'n_cells': len(re.findall(r'<w:tc(?:\s|>)', tr_xml)),
                'has_vmerge': '<w:vMerge' in tr_xml,
                'first_cell_v_align': None,
                'first_cell_n_paras': 0,
                'first_cell_text': '',
            }
            trh_m = re.search(r'<w:trHeight\s+([^/]+?)/>', tr_xml)
            if trh_m:
                attrs = trh_m.group(1)
                vm = re.search(r'w:val="(\d+)"', attrs)
                rm = re.search(r'w:hRule="(\w+)"', attrs)
                tr_meta['trH_val'] = int(vm.group(1)) if vm else None
                tr_meta['h_rule'] = rm.group(1) if rm else None
            # First cell
            tc_m = re.search(r'<w:tc[\s>]', tr_xml)
            if tc_m:
                tc_start = tc_m.start()
                tc_end = tr_xml.find('</w:tc>', tc_start)
                tc_xml = tr_xml[tc_start:tc_end]
                va_m = re.search(r'<w:vAlign\s+w:val="(\w+)"', tc_xml)
                if va_m:
                    tr_meta['first_cell_v_align'] = va_m.group(1)
                tr_meta['first_cell_n_paras'] = len(re.findall(r'<w:p(?:\s|>)', tc_xml))
                txt_m = re.search(r'<w:t[^>]*>([^<]+)</w:t>', tc_xml)
                if txt_m:
                    tr_meta['first_cell_text'] = txt_m.group(1)[:30]
            result[(t_idx, r_idx)] = tr_meta
            r_idx += 1
            rpos = re_pos + len('</w:tr>')
        t_idx += 1
        pos = te + len('</w:tbl>')
    return result


def measure_word_per_row(word, docx_path):
    """Word: per-(table_idx, row_idx) → (y_first_cell, pitch_to_next)."""
    doc = word.Documents.Open(str(docx_path.absolute()), ReadOnly=True)
    out = {}
    try:
        for t in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(t)
            ys = []
            for r in range(1, tbl.Rows.Count + 1):
                try:
                    cell = tbl.Cell(Row=r, Column=1)
                    y = doc.Range(cell.Range.Start, cell.Range.Start).Information(6)
                    ys.append(y)
                except Exception:
                    ys.append(None)
            for r in range(len(ys)):
                pitch = ys[r + 1] - ys[r] if r + 1 < len(ys) and ys[r] and ys[r + 1] else None
                out[(t - 1, r)] = {'y': ys[r], 'pitch': pitch}
    finally:
        doc.Close(SaveChanges=False)
    return out


def measure_oxi_per_row(docx_path):
    """Oxi: per-(cell_row_idx) → first text y. (Single-table assumption per page)."""
    with tempfile.TemporaryDirectory() as tmp:
        prefix = os.path.join(tmp, 'p_')
        dump = os.path.join(tmp, 'layout.json')
        subprocess.run([str(RENDERER), str(docx_path), prefix, '--dump-layout=' + dump],
                      capture_output=True, text=True, timeout=180)
        with open(dump, encoding='utf-8') as f:
            d = json.load(f)
    out = {}  # (table_idx_estimate, row_idx) → y
    # Note: Oxi dumps cell_row_idx without distinguishing tables.
    # Assumption: rows numbered sequentially per page, table boundaries inferred
    # by gaps. Simple version: just collect (row_idx, y) per page.
    for page_idx, page in enumerate(d.get('pages', [])):
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            cr = el.get('cell_row_idx')
            cc = el.get('cell_col_idx')
            if cr is None or cc != 0:
                continue
            key = (page_idx, cr)
            if key not in out or el['y'] < out[key]:
                out[key] = el['y']
    return out


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    try:
        for name in TARGETS:
            docx = DOCX_DIR / name
            if not docx.exists():
                print(f"Skip {name}: not found")
                continue
            print(f"\n=== {name} ===")
            xml_meta = parse_rows_xml(docx)
            word_data = measure_word_per_row(word, docx)
            oxi_data = measure_oxi_per_row(docx)
            print(f"  Tables in XML: {len(set(k[0] for k in xml_meta))}")
            print(f"  Rows in XML: {len(xml_meta)}")
            print(f"  Word rows measured: {len(word_data)}")
            print(f"  Oxi rows measured: {len(oxi_data)}")

            # Compute per-row diffs (limit to T0 = first table)
            # Word data is per (table_idx, row_idx); Oxi is per (page_idx, cell_row_idx).
            # Oxi cell_row_idx resets per table in some configs. Hard to match perfectly.
            # Just dump both sides for offline analysis.
            doc_result = {
                'doc': name,
                'xml_meta_t0': {f'{k[0]},{k[1]}': v for k, v in xml_meta.items() if k[0] == 0},
                'word_t0_rows': {f'{k[0]},{k[1]}': v for k, v in word_data.items() if k[0] == 0},
                'oxi_p0_rows': {f'{k[0]},{k[1]}': v for k, v in oxi_data.items() if k[0] == 0},
            }
            results.append(doc_result)

            # Print summary for T0
            print(f"\n  T0 per-row (Word vs Oxi pitches):")
            for r in range(8):
                xm = xml_meta.get((0, r))
                wd = word_data.get((0, r))
                if not xm:
                    continue
                wp = wd['pitch'] if wd and wd.get('pitch') else None
                # Oxi side: harder to map but for page 0 take cell_row_idx=r
                oxi = oxi_data.get((0, r))
                hrule = xm['h_rule'] or 'auto(default)' if xm.get('trH_val') else '-'
                trh = f"{xm['trH_val']/20:.1f}pt" if xm.get('trH_val') else 'none'
                print(f"    row {r}: hRule={hrule:<15} trH={trh:<10} W_pitch={wp!s:<8} W_y={wd['y'] if wd else None!s:<7} Oxi_y={oxi if oxi else None!s:<7} txt={xm.get('first_cell_text','')!r}")
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT}")


if __name__ == '__main__':
    main()
