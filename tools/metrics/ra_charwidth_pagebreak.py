"""Ra manual measurements: char_width, page_break, spacing collapse"""
import win32com.client, json, os, sys, ctypes
from ctypes import wintypes

results = []

# === 1. CHAR WIDTH: Font fallback for CJK chars with Latin fonts ===
gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32
hdc = user32.GetDC(0)

latin_fonts = ['Calibri', 'Arial', 'Times New Roman', 'Century']
cjk_fonts = ['MS UI Gothic', 'MS Gothic']
test_chars = [
    (0x3042, 'hiragana_a'),
    (0x30A2, 'katakana_a'),
    (0x4E00, 'kanji_one'),
    (0x6587, 'kanji_bun'),
    (0xFF21, 'fullwidth_A'),
    (0x3001, 'comma'),
    (0x3002, 'period'),
    (0x0041, 'latin_A'),
    (0x0061, 'latin_a'),
    (0x0020, 'space'),
]

for font_name in latin_fonts + cjk_fonts:
    charset = 128 if font_name.startswith('MS') else 0
    for ppem in [14, 16, 19]:
        hfont = gdi32.CreateFontW(-ppem, 0, 0, 0, 400, 0, 0, 0, charset, 0, 0, 0, 0, font_name)
        old = gdi32.SelectObject(hdc, hfont)
        for cp, label in test_chars:
            size = wintypes.SIZE()
            gdi32.GetTextExtentPoint32W(hdc, chr(cp), 1, ctypes.byref(size))
            results.append({
                'parameter': f'cw_{font_name}_{label}_ppem{ppem}',
                'value': size.cx,
                'unit': 'px',
                'domain': 'char_width',
                'context': f'font={font_name} char=U+{cp:04X} ppem={ppem}'
            })
        gdi32.SelectObject(hdc, old)
        gdi32.DeleteObject(hfont)

user32.ReleaseDC(0, hdc)
print(f'[1/3] GDI char width: {len(results)} measurements', file=sys.stderr)

# === 2. PAGE BREAK ===
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

try:
    doc = word.Documents.Add()
    sec = doc.Sections(1)
    sec.PageSetup.TopMargin = 72
    sec.PageSetup.BottomMargin = 72
    sec.PageSetup.PageHeight = 841.9
    sec.PageSetup.PageWidth = 595.3

    # Fill page with lines
    for i in range(45):
        if i == 0:
            p = doc.Paragraphs(1)
        else:
            p = doc.Paragraphs.Add()
        p.Range.InsertBefore(f'Line {i+1} test text here')
        p.Range.Font.Name = 'Calibri'
        p.Range.Font.Size = 11
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 0

    # Find first paragraph on page 2
    first_p2 = None
    for i in range(1, doc.Paragraphs.Count + 1):
        page = doc.Paragraphs(i).Range.Information(3)
        if page == 2:
            first_p2 = i
            results.append({
                'parameter': 'first_p2_index',
                'value': i,
                'unit': 'index',
                'domain': 'page_break',
                'context': 'Calibri 11pt Single, 45 lines'
            })
            break

    if first_p2:
        # spaceBefore at page top
        p2 = doc.Paragraphs(first_p2)
        y_no_sb = p2.Range.Information(6)
        p2.Format.SpaceBefore = 24
        y_with_sb = p2.Range.Information(6)
        results.append({
            'parameter': 'page_top_sb_suppressed',
            'value': round(y_with_sb - y_no_sb, 2),
            'unit': 'pt',
            'domain': 'page_break',
            'context': f'sb=24 on first para of page 2: y_no={y_no_sb:.2f} y_with={y_with_sb:.2f}'
        })
        p2.Format.SpaceBefore = 0

        # Widow control
        last_p1 = first_p2 - 1
        doc.Paragraphs(last_p1).Format.WidowControl = True
        page_wc = doc.Paragraphs(last_p1).Range.Information(3)
        results.append({
            'parameter': 'widow_control',
            'value': page_wc,
            'unit': 'page',
            'domain': 'page_break',
            'context': f'WidowControl=True on para {last_p1}'
        })
        doc.Paragraphs(last_p1).Format.WidowControl = False

        # KeepWithNext
        doc.Paragraphs(last_p1 - 1).Format.KeepWithNext = True
        page_kn = doc.Paragraphs(last_p1 - 1).Range.Information(3)
        results.append({
            'parameter': 'keep_with_next',
            'value': page_kn,
            'unit': 'page',
            'domain': 'page_break',
            'context': f'KeepWithNext=True on para {last_p1-1}'
        })
        doc.Paragraphs(last_p1 - 1).Format.KeepWithNext = False

        # KeepTogether (multi-line paragraph)
        # Replace last para on p1 with long text
        doc.Paragraphs(last_p1).Range.Text = 'A very long paragraph ' * 20
        doc.Paragraphs(last_p1).Format.KeepTogether = True
        page_kt = doc.Paragraphs(last_p1).Range.Information(3)
        results.append({
            'parameter': 'keep_together',
            'value': page_kt,
            'unit': 'page',
            'domain': 'page_break',
            'context': f'KeepTogether=True on long para {last_p1}'
        })

    doc.Close(False)
    print(f'[2/3] Page break: done', file=sys.stderr)
except Exception as e:
    print(f'[2/3] Page break error: {e}', file=sys.stderr)

# === 3. SPACING COLLAPSE ===
try:
    doc = word.Documents.Add()
    sec = doc.Sections(1)
    sec.PageSetup.TopMargin = 72

    def setup_pair(font, size, sa, sb):
        p1 = doc.Paragraphs(1)
        p1.Range.Text = 'AAA'
        p1.Range.Font.Name = font
        p1.Range.Font.Size = size
        p1.Format.SpaceBefore = 0
        p1.Format.SpaceAfter = sa
        p1.Format.LineSpacingRule = 0

        if doc.Paragraphs.Count < 2:
            doc.Paragraphs.Add()
        p2 = doc.Paragraphs(2)
        p2.Range.InsertBefore('BBB')
        p2.Range.Font.Name = font
        p2.Range.Font.Size = size
        p2.Format.SpaceBefore = sb
        p2.Format.SpaceAfter = 0
        p2.Format.LineSpacingRule = 0

        y1 = doc.Paragraphs(1).Range.Information(6)
        y2 = doc.Paragraphs(2).Range.Information(6)
        return round(y2 - y1, 2)

    # Get baseline line height (no spacing)
    lh = setup_pair('Calibri', 11, 0, 0)
    results.append({
        'parameter': 'baseline_lh_Calibri_11',
        'value': lh,
        'unit': 'pt',
        'domain': 'spacing',
        'context': 'Calibri 11pt Single no spacing'
    })

    # Test collapse patterns
    tests = [
        (10, 24, 'sa10_sb24'),
        (24, 10, 'sa24_sb10'),
        (10, 10, 'sa10_sb10'),
        (0, 24, 'sa0_sb24'),
        (24, 0, 'sa24_sb0'),
        (15, 15, 'sa15_sb15'),
        (5, 20, 'sa5_sb20'),
    ]

    for sa, sb, label in tests:
        gap = setup_pair('Calibri', 11, sa, sb)
        spacing_total = round(gap - lh, 2)
        results.append({
            'parameter': f'collapse_{label}',
            'value': gap,
            'unit': 'pt',
            'domain': 'spacing',
            'context': f'sa={sa} sb={sb} gap={gap} lh={lh} spacing={spacing_total} (additive={sa+sb} max={max(sa,sb)})'
        })

    # Same tests with MS Gothic
    lh_msg = setup_pair('MS Gothic', 10.5, 0, 0)
    results.append({
        'parameter': 'baseline_lh_MSGothic_10.5',
        'value': lh_msg,
        'unit': 'pt',
        'domain': 'spacing',
        'context': 'MS Gothic 10.5pt Single no spacing'
    })

    for sa, sb, label in [(10, 24, 'sa10_sb24'), (24, 10, 'sa24_sb10'), (10, 10, 'sa10_sb10')]:
        gap = setup_pair('MS Gothic', 10.5, sa, sb)
        spacing_total = round(gap - lh_msg, 2)
        results.append({
            'parameter': f'collapse_MSG_{label}',
            'value': gap,
            'unit': 'pt',
            'domain': 'spacing',
            'context': f'MS Gothic sa={sa} sb={sb} gap={gap} lh={lh_msg} spacing={spacing_total}'
        })

    doc.Close(False)
    print(f'[3/3] Spacing collapse: done', file=sys.stderr)
except Exception as e:
    print(f'[3/3] Spacing error: {e}', file=sys.stderr)

word.Quit()

# Output
print(f'\n=== RESULTS ({len(results)} measurements) ===')
for r in results:
    if r['domain'] != 'char_width':  # Skip verbose char widths
        print(f"  {r['parameter']:50s} = {r['value']} {r['unit']}  ({r['context'][:60]})")

with open('pipeline_data/ra_manual_measurements.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f'\nSaved to pipeline_data/ra_manual_measurements.json', file=sys.stderr)
