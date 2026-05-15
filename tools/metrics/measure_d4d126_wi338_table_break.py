"""COM measurement for d4d126 wi=336..342 (table 4 page-break investigation).

Question: why does Word push the entire single-row 25-paragraph table 4
(containing "匿名データの利用に当たって") to page 7, while Oxi fits its
first content paragraph on page 6?

Output: per-paragraph (page, y, line_spacing, sb, sa, font_size, text) for
the relevant range.
"""
import json, sys, os
import win32com.client
from win32com.client import constants

DOCX = os.path.abspath('tools/golden-test/documents/docx/d4d126dfe1d9_tokumei_08_01-3.docx')
OUT  = os.path.abspath('pipeline_data/d4d126_wi338_table_break.json')

WI_RANGE = list(range(330, 350))

def main():
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        n = doc.Paragraphs.Count
        print(f'Total paragraphs: {n}')
        rows = []
        for wi in WI_RANGE:
            if wi < 1 or wi > n:
                continue
            p = doc.Paragraphs(wi)
            rng = p.Range
            # Collapsed start range for accurate Information(3/5/6)
            start_rng = doc.Range(rng.Start, rng.Start)
            page = start_rng.Information(3)   # wdActiveEndPageNumber on collapsed start
            y    = start_rng.Information(6)   # wdVerticalPositionRelativeToPage
            x    = start_rng.Information(5)   # wdHorizontalPositionRelativeToPage
            line_spacing = p.LineSpacing
            line_spacing_rule = p.LineSpacingRule
            sb = p.SpaceBefore
            sa = p.SpaceAfter
            text = rng.Text or ''
            text = text.replace('\r','').replace('\x07','')[:60]
            # font size: pick first run's font size
            try:
                fs = rng.Font.Size
            except Exception:
                fs = None
            row = dict(
                wi=wi, page=int(page), x=float(x), y=float(y),
                line_spacing=float(line_spacing), line_spacing_rule=int(line_spacing_rule),
                sb=float(sb), sa=float(sa), font_size=float(fs) if fs else None,
                text=text,
            )
            print(f"wi={wi}: page={row['page']} y={row['y']:.2f} ls={row['line_spacing']:.2f}(rule={row['line_spacing_rule']}) sb={row['sb']:.2f} sa={row['sa']:.2f} fs={row['font_size']} text={text!r}")
            rows.append(row)
        with open(OUT, 'w', encoding='utf-8') as f:
            json.dump({'doc': os.path.basename(DOCX), 'rows': rows}, f, ensure_ascii=False, indent=2)
        print(f'\nWrote {OUT}')
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    main()
