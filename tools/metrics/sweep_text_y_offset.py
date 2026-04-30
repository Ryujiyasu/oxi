"""Systematic Word text_y_offset sweep — measure actual offset for
(font × size × grid_pitch) matrix.

Background: R109/R110/R112 hypothesised that Word's grid-snapped
centering uses natural line height (font × CJK 83/64) where Oxi uses
font_size. Single-doc measurements (a5ccbe wp1: 11pt MS Mincho in 18pt
grid → Word offset 2.1pt) supported the hypothesis but broke 3 other
docs when implemented (R112).

Conclusion: 1-doc COM measurement is insufficient. Need systematic
sweep across (font, size, pitch) to derive the formula or a lookup
table.

This script:
1. Generates minimal docx files for each (font, size, pitch) combo
2. Each doc has 3 paragraphs: anchor1, test, anchor2 (all single CJK
   char, anchor1/2 at small fixed size, test at swept (font, size))
3. Opens each in Word via COM, measures Y of each paragraph
4. Records (font, size, pitch) → test_offset_within_grid_cell
5. Saves to pipeline_data/text_y_offset_sweep.json

Output is a data table the layout code can consult or analyse to
derive the right formula.
"""
import os
import sys
import json
import zipfile
import time

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX_DIR = os.path.abspath(os.path.join(
    os.path.dirname(__file__), '..', 'fixtures', 'text_y_offset_sweep'))
RESULTS_PATH = os.path.abspath(os.path.join(
    os.path.dirname(__file__), '..', '..', 'pipeline_data', 'text_y_offset_sweep.json'))

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''


def styles_xml(default_font: str = 'ＭＳ 明朝') -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="Century" w:eastAsia="{default_font}" w:hAnsi="Century" w:cs="Times New Roman"/>'
            '<w:sz w:val="21"/><w:szCs w:val="21"/>'
            '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '</w:styles>')


def build_doc_xml(font_name: str, font_size_halfpt: int, line_pitch_twips: int) -> str:
    """3-paragraph doc:
    - P1 anchor: 'ア' at sz=16 (8pt CJK) — small, consistent baseline
    - P2 test:   'あ' at (font_name, font_size_halfpt)
    - P3 anchor: 'イ' at sz=16
    All paragraphs default snap=true. Section docGrid linePitch varies.
    """
    anchor_pre = (
        '<w:p><w:pPr></w:pPr>'
        '<w:r><w:rPr>'
        '<w:rFonts w:hint="eastAsia"/>'
        '<w:sz w:val="16"/>'
        '</w:rPr>'
        '<w:t>ア</w:t></w:r></w:p>'
    )
    test_para = (
        '<w:p><w:pPr></w:pPr>'
        '<w:r><w:rPr>'
        f'<w:rFonts w:ascii="{font_name}" w:eastAsia="{font_name}" w:hAnsi="{font_name}" w:hint="eastAsia"/>'
        f'<w:sz w:val="{font_size_halfpt}"/>'
        f'<w:szCs w:val="{font_size_halfpt}"/>'
        '</w:rPr>'
        '<w:t>あ</w:t></w:r></w:p>'
    )
    anchor_post = (
        '<w:p><w:pPr></w:pPr>'
        '<w:r><w:rPr>'
        '<w:rFonts w:hint="eastAsia"/>'
        '<w:sz w:val="16"/>'
        '</w:rPr>'
        '<w:t>イ</w:t></w:r></w:p>'
    )
    sect_pr = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1417" w:right="1133" w:bottom="1417" w:left="1133"'
        ' w:header="0" w:footer="0" w:gutter="0"/>'
        '<w:cols w:space="425"/>'
        f'<w:docGrid w:type="lines" w:linePitch="{line_pitch_twips}"/>'
        '</w:sectPr>'
    )
    body = anchor_pre + test_para + anchor_post + sect_pr
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {NS}>'
            f'<w:body>{body}</w:body>'
            '</w:document>')


def make_docx(path: str, font_name: str, font_size_halfpt: int,
              line_pitch_twips: int):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', styles_xml(font_name))
        z.writestr('word/document.xml',
                   build_doc_xml(font_name, font_size_halfpt, line_pitch_twips))


# ---------------------------------------------------------------------------
# Sweep dimensions
# ---------------------------------------------------------------------------
FONTS = [
    'ＭＳ 明朝',     # MS Mincho — bitmapped CJK 83/64
    'ＭＳ ゴシック', # MS Gothic — same family
    '游明朝',        # Yu Mincho — proportional with different metrics
    'Meiryo',        # Meiryo
]

# Sizes in half-points (Word OOXML format)
# 16=8pt, 20=10pt, 21=10.5pt, 22=11pt, 24=12pt, 28=14pt, 32=16pt, 36=18pt, 40=20pt, 48=24pt
SIZES_HALFPT = [16, 20, 21, 22, 24, 28, 36, 40, 48]

# Grid pitches in twips. 240=12pt, 280=14pt, 320=16pt, 360=18pt, 400=20pt
PITCHES_TW = [240, 320, 360, 400]


def measure_via_com(docs):
    """Return {name: {anchor1_y, test_y, anchor2_y}}."""
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    results = {}
    try:
        for name, path in docs:
            try:
                doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
                try:
                    ys = []
                    for i in range(1, 4):
                        para = doc.Paragraphs(i)
                        rng = para.Range
                        rng_start = doc.Range(rng.Start, rng.Start)
                        y = rng_start.Information(6)
                        ys.append(y)
                    results[name] = {
                        'anchor1_y': ys[0],
                        'test_y': ys[1],
                        'anchor2_y': ys[2],
                        'stride_p1_p2': ys[1] - ys[0],
                        'stride_p2_p3': ys[2] - ys[1],
                    }
                finally:
                    doc.Close(SaveChanges=False)
            except Exception as e:
                results[name] = {'error': f'{type(e).__name__}: {e}'}
    finally:
        word.Quit()
    return results


def main(limit: int = 0):
    docs = []
    for font in FONTS:
        for size_hp in SIZES_HALFPT:
            for pitch_tw in PITCHES_TW:
                # Compose name
                font_short = (font.replace('ＭＳ ', 'MS')
                                  .replace('明朝', 'Mincho')
                                  .replace('ゴシック', 'Gothic')
                                  .replace('游', 'Yu')
                                  .replace('Meiryo', 'Meiryo'))
                name = f'{font_short}_sz{size_hp}_pitch{pitch_tw}'
                path = os.path.join(DOCX_DIR, f'{name}.docx')
                make_docx(path, font, size_hp, pitch_tw)
                docs.append((name, path))
                if limit and len(docs) >= limit:
                    break
            if limit and len(docs) >= limit:
                break
        if limit and len(docs) >= limit:
            break

    print(f'Generated {len(docs)} docs')
    print('Measuring via Word COM...')
    t0 = time.time()
    measurements = measure_via_com(docs)
    elapsed = time.time() - t0
    print(f'COM measurement done ({elapsed:.1f}s)')

    # Combine config + measurement
    out = {}
    for font in FONTS:
        for size_hp in SIZES_HALFPT:
            for pitch_tw in PITCHES_TW:
                font_short = (font.replace('ＭＳ ', 'MS')
                                  .replace('明朝', 'Mincho')
                                  .replace('ゴシック', 'Gothic')
                                  .replace('游', 'Yu'))
                name = f'{font_short}_sz{size_hp}_pitch{pitch_tw}'
                if name not in measurements:
                    continue
                m = measurements[name]
                if 'error' in m:
                    out[name] = {'config': {'font': font, 'size_hp': size_hp,
                                            'pitch_tw': pitch_tw},
                                 'error': m['error']}
                    continue
                # Compute test_y_offset within its grid cell.
                # P2 test paragraph starts at cursor_y after anchor.
                # In a docGrid section with snap=true, each paragraph
                # advances by 1+ grid cells. The test paragraph's
                # rendered top (= test_y) sits within its allocated
                # grid cell; offset = test_y - cell_top.
                pitch_pt = pitch_tw / 20.0
                # Anchor1 starts at top_margin (or top + 1 cell). Use
                # anchor1_y as origin.
                # cell_top for test paragraph = anchor1_y + (cells of anchor1)
                # For 8pt anchor in pitch=16pt cell, cells_anchor1 = 1.
                # For test at varied size, cells_test depends on size vs pitch.
                size_pt = size_hp / 2.0
                # Allocated cells: ceil(natural_size / pitch). But Word's
                # actual stride P1→P2 IS the cells × pitch advance.
                stride_p1_p2 = m['stride_p1_p2']
                # Number of cells used by anchor1 = stride / pitch (approx).
                # All paragraphs share same anchor offset within cell when
                # snap=true. So:
                #   anchor1_offset = anchor1_y - anchor1_cell_top
                #   test_offset   = test_y - test_cell_top
                # Without knowing absolute cell_top, compute offset as:
                #   test_y modulo pitch, relative to anchor1_y modulo pitch.
                # Actually the simplest comparison: test_offset_relative =
                #   test_y - anchor1_y - (cells_used_by_anchor1 × pitch)
                # For a single-line anchor at sz=16 (8pt) in pitch≥12pt,
                # cells_used = 1. So:
                cells_anchor1 = max(1, round(stride_p1_p2 / pitch_pt))
                test_cell_top = m['anchor1_y'] + cells_anchor1 * pitch_pt - pitch_pt + pitch_pt
                # Simpler: test_y - anchor1_y = (cells × pitch) + (test_offset - anchor1_offset)
                # If anchor1_offset == test_offset (both rendered with same
                # rule), stride_p1_p2 = cells × pitch exactly. Deviation
                # tells us they differ.
                # For initial analysis just record the y values; compute
                # offset interpretation in a separate analysis script.
                out[name] = {
                    'config': {
                        'font': font,
                        'size_hp': size_hp,
                        'size_pt': size_pt,
                        'pitch_tw': pitch_tw,
                        'pitch_pt': pitch_pt,
                    },
                    'measure': m,
                }

    os.makedirs(os.path.dirname(RESULTS_PATH), exist_ok=True)
    with open(RESULTS_PATH, 'w', encoding='utf-8') as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
    print(f'Saved {len(out)} entries to {RESULTS_PATH}')


if __name__ == '__main__':
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    main(limit=limit)
