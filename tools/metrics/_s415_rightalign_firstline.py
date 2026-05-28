"""S415: COM-measure Word's positioning of the right-aligned + firstLine-indent
cell that S414 identified as the lone ~40pt-off fire cell.

Target: 1ec1091177b1_006.docx, the cell whose paragraph text is
"　　　　税" (4 ideographic spaces + 税), jc=right, firstLineChars=200
(firstLine=420 twip = 21pt). S414 geometry:
  cell content = [322.9, 405.9] pt (page_left 42.55 + col-start 275.4,
  gridSpan=2 width 92.9, pad 4.95)
  Word reports the paragraph start x = 316.0 (LEFT of content-left → overflow)
  Oxi renders text-start x = 356.45, width 52.5 (5 chars × 10.5pt)

Questions to answer with COM:
  1. Per-character horizontal position (Information(5)) of each of the 5
     glyphs → exact glyph rects, so we know where Word puts each char.
  2. Does the firstLine indent (21pt) apply on this right-aligned line?
  3. Are the 4 leading U+3000 rendered full-width, collapsed, or justified?
  4. Is the right edge at content_right (405.9) or cell_right (410.85)?

Also measures an ed025 × × × right-aligned fire cell for cross-check.

Instrumentation only — does NOT modify oxidocs-core or any baseline.
"""
from __future__ import annotations
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')

REPO = r'c:\Users\ryuji\oxi-main'
DOC_1EC1 = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', '1ec1091177b1_006.docx')
DOC_ED025 = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', 'ed025cbecffb_index-23.docx')
OUT = os.path.join(REPO, 'pipeline_data', 'ra_manual_measurements', 's415_rightalign_firstline.json')

# Word Information constants
wdHorizontalPositionRelativeToPage = 5
wdVerticalPositionRelativeToPage = 6
wdActiveEndPageNumber = 3


def measure_paragraph_chars(doc, p):
    """Return per-char (char, x, y) using collapsed single-char ranges."""
    rng = p.Range
    txt = rng.Text or ''
    # Strip trailing cell/para markers
    clean = txt.rstrip('\r\n\x07')
    chars = []
    start = rng.Start
    for i, ch in enumerate(clean):
        cpos = start + i
        try:
            crng = doc.Range(cpos, cpos)  # collapsed to char start
            x = float(crng.Information(wdHorizontalPositionRelativeToPage))
            y = float(crng.Information(wdVerticalPositionRelativeToPage))
        except Exception as e:
            x = y = -1.0
        chars.append({'i': i, 'char': ch, 'cp': hex(ord(ch)), 'x': round(x, 2), 'y': round(y, 2)})
    # Also the position just past the last char (line end)
    try:
        end_rng = doc.Range(start + len(clean), start + len(clean))
        end_x = float(end_rng.Information(wdHorizontalPositionRelativeToPage))
    except Exception:
        end_x = -1.0
    return clean, chars, round(end_x, 2)


def measure_1ec1():
    import win32com.client as wc
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.ScreenUpdating = False
    doc = None
    result = {}
    try:
        doc = word.Documents.Open(DOC_1EC1, ReadOnly=True)
        doc.Repaginate()
        # Find the paragraph "　　　　税"
        target = None
        for pi in range(1, doc.Paragraphs.Count + 1):
            t = (doc.Paragraphs(pi).Range.Text or '').rstrip('\r\n\x07')
            if t.startswith('　　　　') and '税' in t:
                target = doc.Paragraphs(pi)
                break
        if target is None:
            # fallback: any para with 4 leading ideographic spaces
            for pi in range(1, doc.Paragraphs.Count + 1):
                t = (doc.Paragraphs(pi).Range.Text or '')
                if t.count('　') >= 4:
                    target = doc.Paragraphs(pi)
                    print(f'fallback target pi={pi} text={t.rstrip()[:20]!r}')
                    break
        if target is None:
            raise RuntimeError('1ec1 target paragraph not found')

        pf = target.Format
        alignment = pf.Alignment  # 0=left 1=center 2=right 3=justify
        left_indent = float(target.LeftIndent)
        first_line_indent = float(target.FirstLineIndent)
        right_indent = float(target.RightIndent)
        clean, chars, end_x = measure_paragraph_chars(doc, target)
        result = {
            'doc': '1ec1',
            'text': clean,
            'alignment': alignment,
            'left_indent_pt': round(left_indent, 2),
            'first_line_indent_pt': round(first_line_indent, 2),
            'right_indent_pt': round(right_indent, 2),
            'chars': chars,
            'line_end_x': end_x,
        }
        # Cell width context: get the cell containing the paragraph
        try:
            cells = target.Range.Cells
            if cells.Count > 0:
                cell = cells(1)
                result['cell_width_pt'] = round(float(cell.Width), 2)
                result['cell_left_x'] = '(see Information on cell range)'
        except Exception as e:
            result['cell_err'] = str(e)
    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        word.Quit()
    return result


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    out = {}
    print('=== Measuring 1ec1 right-aligned firstLine cell ===')
    r = measure_1ec1()
    out['1ec1'] = r
    print(f"text={r.get('text')!r}")
    print(f"alignment={r.get('alignment')} (2=right) left_indent={r.get('left_indent_pt')} firstLine={r.get('first_line_indent_pt')} right_indent={r.get('right_indent_pt')}")
    print(f"cell_width={r.get('cell_width_pt')}")
    print('per-char positions:')
    for c in r.get('chars', []):
        print(f"  [{c['i']}] {c['cp']} x={c['x']} y={c['y']}")
    print(f"line_end_x={r.get('line_end_x')}")

    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f'\nsaved -> {OUT}')


if __name__ == '__main__':
    main()
