"""Measure Word's actual per-paragraph pitch in ed025c 損益計算書 cells.

Day 37 multi-session investigation. Per [[session58_day37_word_per_row_pitch_17_5pt]],
ed025c PASS gap is Word using ~17.5pt per-paragraph pitch in cell row=2 vs
Oxi's 18pt. This script measures pitch between consecutive paragraphs in
each cell column of the 損益計算書 row, also extracts OOXML pPr/rPr per
paragraph so we can identify spacing source (line, lineRule, etc.).

Output: pipeline_data/ra_manual_measurements/ed025c_cell_pitch.json
        + per-cell pitch histogram printed to stdout.
"""
from __future__ import annotations
import os, sys, json, traceback
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOC = r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\ed025cbecffb_index-23.docx'
OUT = r'c:\Users\ryuji\oxi-main\pipeline_data\ra_manual_measurements\ed025c_cell_pitch.json'


def main():
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.ScreenUpdating = False
    doc = None
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True)
        doc.Repaginate()

        # Find 損益計算書 table via "Ⅰ営業損益"
        target_para_idx = None
        for pi in range(1, doc.Paragraphs.Count + 1):
            txt = (doc.Paragraphs(pi).Range.Text or '').strip()
            if 'Ⅰ' in txt and '営業損益' in txt:
                target_para_idx = pi
                break
        if not target_para_idx:
            print('Target not found')
            return
        table = doc.Paragraphs(target_para_idx).Range.Tables(1)
        print(f'Table found. Rows={table.Rows.Count}, Cols={table.Columns.Count}')

        # For each cell in row 2, measure paragraph-by-paragraph pitch
        results = []
        for cell_idx in range(1, table.Range.Cells.Count + 1):
            cell = table.Range.Cells(cell_idx)
            if cell.RowIndex != 2:  # Only row 2 (big content row)
                continue
            col = cell.ColumnIndex
            paras = cell.Range.Paragraphs
            cell_data = {
                'col': col,
                'n_paragraphs': paras.Count,
                'paragraphs': [],
            }
            prev_y = None
            prev_page = None
            for pi in range(1, paras.Count + 1):
                p = paras(pi)
                rng = p.Range
                start = doc.Range(rng.Start, rng.Start)
                try:
                    page = int(start.Information(1))
                    y = float(start.Information(6))
                except:
                    page = -1
                    y = -1.0
                txt = (p.Range.Text or '').rstrip('\r\n\x07')[:25]
                # Read pPr line/lineRule/spaceBefore/spaceAfter
                fmt = p.Format
                line_spacing = float(fmt.LineSpacing) if hasattr(fmt, 'LineSpacing') else 0.0
                line_spacing_rule = int(fmt.LineSpacingRule) if hasattr(fmt, 'LineSpacingRule') else -1
                space_before = float(fmt.SpaceBefore) if hasattr(fmt, 'SpaceBefore') else 0.0
                space_after = float(fmt.SpaceAfter) if hasattr(fmt, 'SpaceAfter') else 0.0
                # Pitch from prev paragraph (same page)
                if prev_y is not None and page == prev_page:
                    pitch = y - prev_y
                else:
                    pitch = None
                cell_data['paragraphs'].append({
                    'pi': pi - 1,
                    'page': page,
                    'y': round(y, 2),
                    'pitch_from_prev': round(pitch, 2) if pitch is not None else None,
                    'ls': line_spacing,
                    'ls_rule': line_spacing_rule,
                    'sb': space_before,
                    'sa': space_after,
                    'text': txt,
                })
                prev_y = y
                prev_page = page
            results.append(cell_data)
            print(f'\nCell col={col}: {paras.Count} paragraphs')
            # Pitch histogram for this cell
            pitches = [p['pitch_from_prev'] for p in cell_data['paragraphs'] if p['pitch_from_prev'] is not None]
            from collections import Counter
            histo = Counter([round(p, 1) for p in pitches])
            print(f'  Pitch histogram: {dict(sorted(histo.items()))}')

        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        with open(OUT, 'w', encoding='utf-8') as f:
            json.dump({'doc': 'ed025c', 'cells': results}, f, ensure_ascii=False, indent=2)
        print(f'\nSaved to {OUT}')

        # Also extract OOXML pPr for first ~10 paragraphs of cell col=3 (the × × × col)
        print('\n=== OOXML pPr/rPr inspection (cell col=3) ===')
        col3 = next((c for c in results if c['col'] == 3), None)
        if col3:
            import zipfile, re
            with zipfile.ZipFile(DOC) as z:
                with z.open('word/document.xml') as f:
                    x = f.read().decode('utf-8')
            # Find cell col=3 in the 損益計算書 table; just count <w:tc> from the table start
            target_loc = x.find('Ⅰ　営業損益')
            tbl_start = x.rfind('<w:tbl>', 0, target_loc)
            tbl_end = x.find('</w:tbl>', target_loc)
            tbl_xml = x[tbl_start:tbl_end]
            # Within the table, find <w:tr> 2 (Row index 2 in Word is the big content row)
            tr_starts = [m.start() for m in re.finditer(r'<w:tr[\s>]', tbl_xml)]
            print(f'TR opens in table: {len(tr_starts)}')
            if len(tr_starts) >= 2:
                tr2_start = tr_starts[1]  # row 2 = index 1 (0-based)
                tr2_end = tbl_xml.find('</w:tr>', tr2_start)
                tr2_xml = tbl_xml[tr2_start:tr2_end]
                # Find tcs
                tc_starts = [m.start() for m in re.finditer(r'<w:tc[\s>]', tr2_xml)]
                print(f'TC opens in row 2: {len(tc_starts)}')
                # Cell col 3 = tc index 2 (0-based)
                if len(tc_starts) >= 3:
                    tc3_start = tc_starts[2]
                    tc3_end = tr2_xml.find('</w:tc>', tc3_start)
                    tc3_xml = tr2_xml[tc3_start:tc3_end]
                    # First 5 paragraphs
                    para_starts = [m.start() for m in re.finditer(r'<w:p[\s>]', tc3_xml)]
                    print(f'Paragraphs in cell col=3: {len(para_starts)}')
                    for i, ps in enumerate(para_starts[:8]):
                        pe = tc3_xml.find('</w:p>', ps)
                        para_xml = tc3_xml[ps:pe]
                        # Find pPr
                        ppr_match = re.search(r'<w:pPr>.*?</w:pPr>', para_xml, re.DOTALL)
                        ppr = ppr_match.group(0) if ppr_match else '(no pPr)'
                        txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para_xml))[:25]
                        print(f'  p{i}: text={txt!r}')
                        print(f'      pPr: {ppr[:300]}')

    except Exception:
        traceback.print_exc()
    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        word.Quit()


if __name__ == '__main__':
    main()
