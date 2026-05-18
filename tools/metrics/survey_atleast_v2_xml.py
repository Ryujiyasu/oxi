"""S105: extended survey reading OOXML hRule attribute directly.

S104 survey used `row.HeightRule` from COM which conflates:
- explicit `<w:trHeight w:val="N" w:hRule="atLeast"/>` → COM HR=1
- no hRule `<w:trHeight w:val="N"/>` (= default "auto" per ECMA-376) → COM HR=1 too

This survey reads the OOXML XML directly to distinguish these,
PLUS captures Word's actual rendering. Goal: find if any class has
clean Word formula.
"""
import json
import re
import sys
import zipfile
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path('c:/Users/ryuji/oxi-main')
DOCX_DIR = ROOT / 'tools/golden-test/documents/docx'
OUT = ROOT / 'tools/metrics/atleast_snap_survey_v2.json'

wdVerticalPositionRelativeToPage = 6


def parse_table_rows_from_xml(docx_path: Path) -> list[dict]:
    """Parse OOXML XML and return per-row trHeight info."""
    with zipfile.ZipFile(docx_path) as z:
        try:
            doc = z.read('word/document.xml').decode('utf-8', errors='ignore')
        except KeyError:
            return []
    # Find tables, iterate rows
    tables = []
    pos = 0
    while True:
        tbl_start = doc.find('<w:tbl>', pos)
        if tbl_start < 0:
            tbl_start = doc.find('<w:tbl ', pos)
            if tbl_start < 0:
                break
        tbl_end = doc.find('</w:tbl>', tbl_start)
        if tbl_end < 0:
            break
        tbl_xml = doc[tbl_start:tbl_end]
        rows = []
        rpos = 0
        while True:
            tr_match = re.search(r'<w:tr(?:\s|>)', tbl_xml[rpos:])
            if not tr_match:
                break
            tr_start = rpos + tr_match.start()
            tr_end = tbl_xml.find('</w:tr>', tr_start)
            if tr_end < 0:
                break
            tr_xml = tbl_xml[tr_start:tr_end]
            # Extract trHeight
            trh_match = re.search(r'<w:trHeight\s+([^/]+?)/>', tr_xml)
            row_info = {'has_trH': trh_match is not None, 'h_rule_xml': None, 'trH_val': None}
            if trh_match:
                attrs = trh_match.group(1)
                val_m = re.search(r'w:val="(\d+)"', attrs)
                rule_m = re.search(r'w:hRule="(\w+)"', attrs)
                row_info['trH_val'] = int(val_m.group(1)) if val_m else None
                row_info['h_rule_xml'] = rule_m.group(1) if rule_m else None  # None = default = "auto" per spec
            rows.append(row_info)
            rpos = tr_end + len('</w:tr>')
        tables.append(rows)
        pos = tbl_end + len('</w:tbl>')
    return tables


def measure_doc(word, docx_path: Path) -> list[dict]:
    """Combine XML-parsed row info with Word COM-measured pitches."""
    xml_tables = parse_table_rows_from_xml(docx_path)
    if not xml_tables:
        return []

    doc = word.Documents.Open(str(docx_path.absolute()), ReadOnly=True)
    rows_data = []
    try:
        for t_idx in range(min(doc.Tables.Count, len(xml_tables))):
            tbl = doc.Tables(t_idx + 1)
            xml_rows = xml_tables[t_idx]
            n_com_rows = tbl.Rows.Count
            for r_idx in range(min(n_com_rows, len(xml_rows))):
                xml_row = xml_rows[r_idx]
                try:
                    com_row = tbl.Rows(r_idx + 1)
                    cell1 = tbl.Cell(Row=r_idx + 1, Column=1)
                    y_this = doc.Range(cell1.Range.Start, cell1.Range.Start).Information(wdVerticalPositionRelativeToPage)
                    pitch = None
                    if r_idx + 1 < n_com_rows:
                        cell_next = tbl.Cell(Row=r_idx + 2, Column=1)
                        y_next = doc.Range(cell_next.Range.Start, cell_next.Range.Start).Information(wdVerticalPositionRelativeToPage)
                        pitch = y_next - y_this
                    h_rule_com = com_row.HeightRule
                    border_pt = tbl.Borders.OutsideLineWidth or 0.5
                    rows_data.append({
                        'doc': docx_path.name,
                        'table_idx': t_idx,
                        'row_idx': r_idx,
                        'h_rule_xml': xml_row['h_rule_xml'],  # None/explicit
                        'h_rule_xml_effective': xml_row['h_rule_xml'] or ('auto' if xml_row['has_trH'] else None),
                        'has_trH_xml': xml_row['has_trH'],
                        'trH_xml_val': xml_row['trH_val'],  # in twip
                        'trH_xml_pt': xml_row['trH_val'] / 20 if xml_row['trH_val'] else None,
                        'h_rule_com': h_rule_com,
                        'rendered_pitch_pt': pitch,
                        'border_pt': border_pt,
                    })
                except Exception:
                    continue
    finally:
        doc.Close(SaveChanges=False)
    return rows_data


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    all_rows = []
    try:
        docs = sorted([p for p in DOCX_DIR.glob('*.docx')
                       if not p.name.startswith('test_')
                       and not p.name.startswith('~')])
        print(f"Surveying {len(docs)} docs...")
        for i, p in enumerate(docs):
            try:
                rows = measure_doc(word, p)
                all_rows.extend(rows)
                if (i + 1) % 20 == 0:
                    print(f"  [{i+1}/{len(docs)}] {len(all_rows)} rows so far")
            except Exception as e:
                print(f"  [{i+1}] {p.name}: ERR {e}")
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(all_rows, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {len(all_rows)} rows to {OUT}")


if __name__ == '__main__':
    main()
