"""Session 130 — COM measurement of vertical-writing cells.

For each input docx, opens in Word and measures:
  - For each table cell with textDirection=tbRlV (or btLr):
    - cell index (table, row, col)
    - cell width/height/x/y
    - vAlign value (from XML)
    - text inside (raw)
    - per-paragraph y inside cell (collapsed-start fix applied)
  - For each adjacent horizontal cell in the same row:
    - per-paragraph y
  - Row height (computed from row top y to next row top y)

Output: pipeline_data/vert_writing_measurements_S130.json

Pre-reads XML to identify cells with textDirection (Word COM does not
expose this property directly), then maps them to (table, row, col) via
indexing. Then opens in Word COM and pulls geometry.

Run with: python tools/metrics/com_measure_vert_writing.py
"""
import json
import os
import re
import sys
import time
import zipfile

sys.stdout.reconfigure(encoding='utf-8')

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_PATH = os.path.join(REPO_ROOT, "pipeline_data", "vert_writing_measurements_S130.json")

REPROS = [
    "tools/golden-test/repros/vert_writing_S130/V1_basic.docx",
    "tools/golden-test/repros/vert_writing_S130/V2_tall_text.docx",
    "tools/golden-test/repros/vert_writing_S130/V3_short_text.docx",
    "tools/golden-test/repros/vert_writing_S130/V4_no_valign.docx",
    "tools/golden-test/repros/vert_writing_S130/V5_top_valign.docx",
]

REAL_DOCS = [
    "tools/golden-test/documents/docx/2ea81a8441cc_0025006-192.docx",
    "tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx",
    "tools/golden-test/documents/docx/7ead52b63f0e_000067058.docx",
    "tools/golden-test/documents/docx/ed025cbecffb_index-23.docx",
]


def find_vert_cells_from_xml(docx_path: str) -> list[dict]:
    """Pre-scan document.xml to locate cells with vertical textDirection.

    Returns list of dicts with absolute byte offsets and contextual info.
    The COM step later indexes Tables/Rows/Cells by document order and
    matches by sequence number.
    """
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read('word/document.xml').decode('utf-8', errors='replace')

    def rfind_tag(haystack: str, tag: str, end: int) -> int:
        """Find the last `<w:{tag}` followed by ` ` or `>` in haystack[:end]."""
        pat = re.compile(r'<w:' + tag + r'[ >]')
        last = -1
        for m in pat.finditer(haystack, 0, end):
            last = m.start()
        return last

    results = []
    # Find each <w:textDirection val="tbRl*"/>
    for m in re.finditer(r'<w:textDirection w:val="(tbRl[VR]?|btLr)"', xml):
        td_val = m.group(1)
        td_pos = m.start()
        # Walk back to find enclosing <w:tc>
        tc_start = rfind_tag(xml, 'tc', td_pos)
        tc_end = xml.find('</w:tc>', td_pos) + len('</w:tc>')
        # Walk back to find enclosing <w:tr>
        tr_start = rfind_tag(xml, 'tr', tc_start)
        # Walk back to find enclosing <w:tbl>
        tbl_start = rfind_tag(xml, 'tbl', tr_start)

        # Compute cell index within row: count <w:tc occurrences from tr_start to tc_start
        tr_to_tc = xml[tr_start:tc_start]
        cell_idx_in_row = len(re.findall(r'<w:tc[ >]', tr_to_tc))  # 0-based

        # Compute row index within table: count <w:tr> from tbl_start to tr_start
        tbl_to_tr = xml[tbl_start:tr_start]
        row_idx_in_tbl = len(re.findall(r'<w:tr[ >]', tbl_to_tr))  # 0-based

        # Compute table index in doc: count <w:tbl> before this one (top-level only — naive)
        tbl_idx = len(re.findall(r'<w:tbl[ >]', xml[:tbl_start]))  # 0-based

        # Extract vAlign
        tc_xml = xml[tc_start:tc_end]
        valign_m = re.search(r'<w:vAlign w:val="(\w+)"', tc_xml)
        valign = valign_m.group(1) if valign_m else None

        # Extract cell width
        tcw_m = re.search(r'<w:tcW w:w="(\d+)" w:type="(\w+)"', tc_xml)
        tcw = (int(tcw_m.group(1)), tcw_m.group(2)) if tcw_m else None

        # Extract text content
        texts = re.findall(r'<w:t[^>]*>([^<]+)</w:t>', tc_xml)
        text = ''.join(texts)

        # Extract sz from rPr (half-points)
        sz_m = re.search(r'<w:sz w:val="(\d+)"', tc_xml)
        sz_hp = int(sz_m.group(1)) if sz_m else None

        results.append({
            'tbl_idx': tbl_idx + 1,  # 1-based for COM
            'row_idx': row_idx_in_tbl + 1,
            'col_idx': cell_idx_in_row + 1,
            'textDirection': td_val,
            'vAlign': valign,
            'tcW_dxa': tcw[0] if tcw else None,
            'tcW_type': tcw[1] if tcw else None,
            'text': text,
            'sz_hp': sz_hp,
        })
    return results


def measure_doc(word_app, docx_path: str) -> dict:
    """Open doc in Word, measure vertical-cell geometry."""
    abs_path = os.path.abspath(docx_path).replace('/', '\\')
    print(f'  Opening: {abs_path}')
    vert_cells = find_vert_cells_from_xml(docx_path)
    print(f'    Found {len(vert_cells)} vertical cell(s) via XML scan')
    if not vert_cells:
        return {'doc': docx_path, 'vert_cells': [], 'note': 'no vertical cells'}

    doc = word_app.Documents.Open(abs_path, ReadOnly=True)
    try:
        measurements = []
        for vc in vert_cells:
            try:
                tbl = doc.Tables(vc['tbl_idx'])
                cell = tbl.Cell(vc['row_idx'], vc['col_idx'])
                rng = cell.Range
                # collapsed start (R30 fix)
                start_rng = doc.Range(rng.Start, rng.Start)
                end_rng = doc.Range(rng.End, rng.End)
                # Information constants:
                #   wdHorizontalPositionRelativeToPage = 5
                #   wdVerticalPositionRelativeToPage = 6
                cell_x = start_rng.Information(5)
                cell_y = start_rng.Information(6)
                cell_w_pt = cell.Width  # points
                # Cell height in points
                try:
                    cell_h_pt = cell.Height
                except Exception:
                    cell_h_pt = None

                # Adjacent cell (col+1 in same row), if exists
                adj_y_paras = []
                try:
                    adj_cell = tbl.Cell(vc['row_idx'], vc['col_idx'] + 1)
                    adj_x = doc.Range(adj_cell.Range.Start, adj_cell.Range.Start).Information(5)
                    adj_y = doc.Range(adj_cell.Range.Start, adj_cell.Range.Start).Information(6)
                    adj_w = adj_cell.Width
                    # Per-paragraph y inside adj cell
                    for i, p in enumerate(adj_cell.Range.Paragraphs):
                        ps = doc.Range(p.Range.Start, p.Range.Start)
                        adj_y_paras.append({
                            'i': i,
                            'y': ps.Information(6),
                            'x': ps.Information(5),
                            'text': p.Range.Text.strip()[:60],
                        })
                except Exception as e:
                    adj_x = None; adj_y = None; adj_w = None

                measurements.append({
                    **vc,
                    'cell_x_pt': cell_x,
                    'cell_y_pt': cell_y,
                    'cell_w_pt': cell_w_pt,
                    'cell_h_pt': cell_h_pt,
                    'cell_end_y_pt': end_rng.Information(6),
                    'adj_x_pt': adj_x,
                    'adj_y_pt': adj_y,
                    'adj_w_pt': adj_w,
                    'adj_paras': adj_y_paras,
                })
            except Exception as e:
                print(f'    ERROR measuring cell {vc}: {e}')
                measurements.append({**vc, 'error': str(e)})
        return {'doc': docx_path, 'vert_cells': measurements}
    finally:
        doc.Close(SaveChanges=False)


def main():
    try:
        import win32com.client as win32
    except ImportError:
        print('ERROR: pywin32 not installed. Run: pip install pywin32')
        sys.exit(1)

    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        results = []
        all_docs = REPROS + REAL_DOCS
        for path in all_docs:
            if not os.path.exists(path):
                print(f'SKIP (missing): {path}')
                continue
            print(f'\n=== {path} ===')
            try:
                r = measure_doc(word, path)
                results.append(r)
            except Exception as e:
                print(f'  FAILED: {e}')
                results.append({'doc': path, 'error': str(e)})

        os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
        with open(OUT_PATH, 'w', encoding='utf-8') as f:
            json.dump({'docs': results, 'generated': time.time()}, f, ensure_ascii=False, indent=2)
        print(f'\nWrote: {OUT_PATH}')
    finally:
        word.Quit()


if __name__ == '__main__':
    main()
