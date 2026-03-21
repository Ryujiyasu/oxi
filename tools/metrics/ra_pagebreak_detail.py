"""Page break detailed measurements"""
import win32com.client, json, sys

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []

def add(param, value, context):
    results.append({
        'parameter': param, 'value': value,
        'unit': 'pt', 'domain': 'page_break', 'context': context
    })

# === 1. Orphan control (last line alone on next page) ===
try:
    doc = word.Documents.Add()
    sec = doc.Sections(1)
    sec.PageSetup.TopMargin = 72
    sec.PageSetup.BottomMargin = 72
    sec.PageSetup.PageHeight = 841.9

    # Fill to near bottom with single-line paragraphs
    for i in range(38):
        if i == 0:
            p = doc.Paragraphs(1)
        else:
            p = doc.Paragraphs.Add()
        p.Range.InsertBefore(f'Line {i+1}')
        p.Range.Font.Name = 'Calibri'
        p.Range.Font.Size = 11
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 0
        p.Format.WidowControl = False

    # Add a 3-line paragraph at the boundary
    p_long = doc.Paragraphs.Add()
    p_long.Range.InsertBefore('A ' * 200)  # wraps to ~3 lines
    p_long.Range.Font.Name = 'Calibri'
    p_long.Range.Font.Size = 11
    p_long.Format.SpaceBefore = 0
    p_long.Format.SpaceAfter = 0
    p_long.Format.LineSpacingRule = 0
    p_long.Format.WidowControl = False

    # Check: which page is the long paragraph on?
    long_page = p_long.Range.Information(3)
    long_y = p_long.Range.Information(6)
    add('orphan_test_no_control', long_page, f'3-line para at boundary, WidowControl=False, page={long_page}, y={long_y:.2f}')

    # Enable widow control
    p_long.Format.WidowControl = True
    long_page_wc = p_long.Range.Information(3)
    long_y_wc = p_long.Range.Information(6)
    add('orphan_test_with_control', long_page_wc, f'WidowControl=True, page={long_page_wc}, y={long_y_wc:.2f}')

    doc.Close(False)
    print('[1/4] Orphan control: done', file=sys.stderr)
except Exception as e:
    print(f'[1/4] Orphan error: {e}', file=sys.stderr)

# === 2. Table row splitting across pages ===
try:
    doc = word.Documents.Add()
    sec = doc.Sections(1)
    sec.PageSetup.TopMargin = 72
    sec.PageSetup.BottomMargin = 72
    sec.PageSetup.PageHeight = 841.9

    # Fill most of page 1
    for i in range(30):
        if i == 0:
            p = doc.Paragraphs(1)
        else:
            p = doc.Paragraphs.Add()
        p.Range.InsertBefore(f'Line {i+1}')
        p.Range.Font.Name = 'Calibri'
        p.Range.Font.Size = 11
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 0

    # Add a table with tall row
    r = doc.Paragraphs(doc.Paragraphs.Count).Range
    tbl = doc.Tables.Add(r, 3, 2)

    # Make row 2 very tall
    for pi in range(1, tbl.Cell(2, 1).Range.Paragraphs.Count + 1):
        tbl.Cell(2, 1).Range.Paragraphs(pi).Range.InsertBefore('Cell text line\r')
    for _ in range(10):
        tbl.Cell(2, 1).Range.InsertAfter('More cell text\r')

    # Check: does the table row split?
    row1_page = tbl.Rows(1).Range.Information(3)
    row2_page = tbl.Rows(2).Range.Information(3)
    row3_page = tbl.Rows(3).Range.Information(3)

    add('table_row1_page', row1_page, 'Table row 1')
    add('table_row2_page', row2_page, 'Table row 2 (tall)')
    add('table_row3_page', row3_page, 'Table row 3')

    # Check AllowBreakAcrossPages
    allow = tbl.Rows(2).AllowBreakAcrossPages
    add('table_allow_break_default', 1 if allow else 0, f'AllowBreakAcrossPages default={allow}')

    # Disable row break
    tbl.Rows(2).AllowBreakAcrossPages = False
    row2_page_no_break = tbl.Rows(2).Range.Information(3)
    add('table_row2_no_break', row2_page_no_break, 'Row 2 with AllowBreak=False')

    doc.Close(False)
    print('[2/4] Table row split: done', file=sys.stderr)
except Exception as e:
    print(f'[2/4] Table error: {e}', file=sys.stderr)

# === 3. SpaceBefore suppression at page/column top ===
try:
    doc = word.Documents.Add()
    sec = doc.Sections(1)
    sec.PageSetup.TopMargin = 72
    sec.PageSetup.BottomMargin = 72

    # Fill page 1
    for i in range(38):
        if i == 0:
            p = doc.Paragraphs(1)
        else:
            p = doc.Paragraphs.Add()
        p.Range.InsertBefore(f'Line {i+1}')
        p.Range.Font.Name = 'Calibri'
        p.Range.Font.Size = 11
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 0

    # Add paragraphs on page 2 with various spaceBefore
    for sb_val in [0, 6, 12, 24, 48]:
        p = doc.Paragraphs.Add()
        p.Range.InsertBefore(f'Page2 sb={sb_val}')
        p.Range.Font.Name = 'Calibri'
        p.Range.Font.Size = 11
        p.Format.SpaceBefore = sb_val
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 0

    # Measure Y positions on page 2
    prev_y = None
    for i in range(39, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        page = p.Range.Information(3)
        y = p.Range.Information(6)
        sb = p.Format.SpaceBefore
        text = p.Range.Text.strip()[:20]
        if page == 2:
            gap = round(y - prev_y, 2) if prev_y else 0
            add(f'page2_sb{int(sb)}_y', y, f'page 2 sb={sb} y={y:.2f} gap_from_prev={gap}')
            prev_y = y

    doc.Close(False)
    print('[3/4] SpaceBefore suppression: done', file=sys.stderr)
except Exception as e:
    print(f'[3/4] SpaceBefore error: {e}', file=sys.stderr)

# === 4. contextualSpacing ===
try:
    doc = word.Documents.Add()
    sec = doc.Sections(1)
    sec.PageSetup.TopMargin = 72

    # Two Normal paragraphs with sa=10
    p1 = doc.Paragraphs(1)
    p1.Range.Text = 'Normal para 1'
    p1.Range.Font.Name = 'Calibri'
    p1.Range.Font.Size = 11
    p1.Style = doc.Styles('Normal')
    p1.Format.SpaceAfter = 10
    p1.Format.SpaceBefore = 0
    p1.Format.LineSpacingRule = 0

    p2 = doc.Paragraphs.Add()
    p2.Range.InsertBefore('Normal para 2')
    p2.Range.Font.Name = 'Calibri'
    p2.Range.Font.Size = 11
    p2.Style = doc.Styles('Normal')
    p2.Format.SpaceAfter = 10
    p2.Format.SpaceBefore = 0
    p2.Format.LineSpacingRule = 0

    y1 = doc.Paragraphs(1).Range.Information(6)
    y2 = doc.Paragraphs(2).Range.Information(6)
    gap_no_ctx = round(y2 - y1, 2)

    # Enable contextualSpacing
    p1.Format.ContextualSpacing = True
    p2.Format.ContextualSpacing = True
    y1c = doc.Paragraphs(1).Range.Information(6)
    y2c = doc.Paragraphs(2).Range.Information(6)
    gap_ctx = round(y2c - y1c, 2)

    add('contextual_spacing_off', gap_no_ctx, f'Two Normal paras sa=10, no contextualSpacing')
    add('contextual_spacing_on', gap_ctx, f'Two Normal paras sa=10, contextualSpacing=True')

    # Different styles: contextualSpacing should NOT apply
    p2.Style = doc.Styles('Heading 1')
    p2.Range.Font.Size = 11
    p2.Format.ContextualSpacing = True
    y2d = doc.Paragraphs(2).Range.Information(6)
    gap_diff = round(y2d - y1c, 2)
    add('contextual_spacing_diff_style', gap_diff, f'Normal+Heading1 with contextualSpacing=True')

    doc.Close(False)
    print('[4/4] ContextualSpacing: done', file=sys.stderr)
except Exception as e:
    print(f'[4/4] ContextualSpacing error: {e}', file=sys.stderr)

word.Quit()

print(f'\n=== RESULTS ({len(results)} measurements) ===')
for r in results:
    print(f"  {r['parameter']:45s} = {r['value']:>8} {r['unit']}  {r['context'][:70]}")

# Append to existing file
try:
    with open('pipeline_data/ra_manual_measurements.json', encoding='utf-8') as f:
        existing = json.load(f)
    existing.extend(results)
    with open('pipeline_data/ra_manual_measurements.json', 'w', encoding='utf-8') as f:
        json.dump(existing, f, indent=2, ensure_ascii=False)
    print(f'\nAppended to ra_manual_measurements.json (total: {len(existing)})', file=sys.stderr)
except:
    with open('pipeline_data/ra_manual_measurements.json', 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
