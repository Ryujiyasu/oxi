"""Day 33 part 32 — Compare db9ca wi=37 vs 0e7af wi=232 paragraph properties.

Both inherit widow_control=false from docDefaults. Yet:
- db9ca wi=37: Word fits line 1 on page 2 (Day 33 part 30 confirmed)
- 0e7af wi=232: Word pushes whole paragraph to page 7

What's the discriminator?
"""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

CASES = [
    ('db9ca', 'tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx', 37),
    ('0e7af', 'tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx', 232),
]


def measure(label, docx, para_idx):
    word = wc.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    abs_path = os.path.abspath(docx)
    d = word.Documents.Open(abs_path, ReadOnly=True)
    try:
        p = d.Paragraphs(para_idx)
        r = p.Range
        text = (r.Text or '').rstrip('\r\n')[:80]
        cr = d.Range(r.Start, r.Start)
        try: pg = int(cr.Information(3))
        except: pg = -1
        try: y = round(cr.Information(6), 2)
        except: y = -1
        try: fs = r.Font.Size
        except: fs = -1
        try: font_name = r.Font.Name
        except: font_name = '?'
        try: lh = p.Format.LineSpacing
        except: lh = -1
        try: lhr = p.Format.LineSpacingRule
        except: lhr = -1
        try: sb = p.Format.SpaceBefore
        except: sb = -1
        try: sa = p.Format.SpaceAfter
        except: sa = -1
        try: snap = p.Format.SnapToGrid
        except: snap = None
        try: keep_with_next = p.Format.KeepWithNext
        except: keep_with_next = None
        try: keep_together = p.Format.KeepTogether
        except: keep_together = None
        try: page_break_before = p.Format.PageBreakBefore
        except: page_break_before = None
        try: widow = p.Format.WidowControl
        except: widow = None
        try: ali = p.Format.Alignment
        except: ali = -1
        try: sty = str(p.Style.NameLocal)
        except: sty = '?'
        try: li = p.Format.LeftIndent
        except: li = 0
        try: ri = p.Format.RightIndent
        except: ri = 0
        try: fli = p.Format.FirstLineIndent
        except: fli = 0
        try: ta = p.Format.TextAlignment
        except: ta = -1
        # Number of lines
        try: n_lines = int(r.Information(8))  # wdNumberOfPagesInDocument - wrong
        except: n_lines = -1
        # Check for LRPB - need to walk runs
        # Range.Text doesn't show LRPB, but we can check OOXML
        char_count = r.Characters.Count
        # Per-line measurement via Characters loop
        line_breaks = []
        prev_y = None
        prev_pg = None
        for i in range(1, char_count + 1):
            c = r.Characters(i)
            ch_range = d.Range(c.Start, c.Start)
            try:
                cpg = int(ch_range.Information(3))
                cy = round(ch_range.Information(6), 2)
            except: continue
            if prev_y is None or cpg != prev_pg or abs(cy - prev_y) > 0.5:
                line_breaks.append((i, cpg, cy))
            prev_y = cy
            prev_pg = cpg
        print(f'=== {label} wi={para_idx} ===')
        print(f'  text: {text!r}')
        print(f'  pg={pg} y={y}')
        print(f'  font: {font_name!r} fs={fs}')
        print(f'  style: {sty!r}')
        print(f'  lh={lh} lhr={lhr}')
        print(f'  sb={sb} sa={sa}')
        print(f'  snap={snap}')
        print(f'  widow={widow}')
        print(f'  keep_next={keep_with_next} keep_together={keep_together}')
        print(f'  pageBreakBefore={page_break_before}')
        print(f'  alignment={ali} textAlignment={ta}')
        print(f'  indent: left={li} right={ri} fli={fli}')
        print(f'  char_count={char_count}')
        print(f'  line_breaks count: {len(line_breaks)}')
        for lb in line_breaks[:6]:
            print(f'    line[char{lb[0]}] pg={lb[1]} y={lb[2]}')
        print()
    finally:
        d.Close(False)
        word.Quit()


if __name__ == '__main__':
    for label, docx, idx in CASES:
        measure(label, docx, idx)
