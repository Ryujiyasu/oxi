"""Measure Word's actual footer top position on ed025c page 13.

To determine if Oxi over-reserves footer area, compare Word's footer text Y
to Oxi's content bottom. If Word's footer text starts at y=X, Word's body
content area = [top_margin, X].
"""
from __future__ import annotations
import os, sys, traceback
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOC = r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\ed025cbecffb_index-23.docx'


def main():
    word = wc.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.ScreenUpdating = False
    doc = None
    try:
        doc = word.Documents.Open(DOC, ReadOnly=True)
        doc.Repaginate()
        # Get section page setup
        section = doc.Sections(1)
        ps = section.PageSetup
        print('Page setup:')
        print(f'  page_width: {ps.PageWidth} pt')
        print(f'  page_height: {ps.PageHeight} pt')
        print(f'  top_margin: {ps.TopMargin} pt')
        print(f'  bottom_margin: {ps.BottomMargin} pt')
        print(f'  footer_distance: {ps.FooterDistance} pt')
        print(f'  header_distance: {ps.HeaderDistance} pt')
        print()
        # Compute Word's effective body content area
        print(f'Word body content bottom:')
        print(f'  = page_height - bottom_margin = {ps.PageHeight - ps.BottomMargin}')
        print(f'  = page_height - footer_distance = {ps.PageHeight - ps.FooterDistance}')
        print()
        # Get footer range and measure its first text y position
        section_footer = section.Footers(2)  # 2 = wdHeaderFooterPrimary
        print(f'Footer range: {section_footer.Range.Start} to {section_footer.Range.End}')
        ft_text = section_footer.Range.Text
        print(f'Footer text repr: {ft_text!r}')

        # Get the page 13 last paragraph y
        # Iterate paragraphs to find ones on page 13
        n_paras = doc.Paragraphs.Count
        last_p13_y = 0
        last_p13_text = ''
        first_p14_y = 999
        first_p14_text = ''
        for pi in range(1, n_paras + 1):
            try:
                para = doc.Paragraphs(pi)
                rng = para.Range
                start_rng = doc.Range(rng.Start, rng.Start)
                page = int(start_rng.Information(1))
                y = float(start_rng.Information(6))
                if page == 13:
                    if y > last_p13_y:
                        last_p13_y = y
                        last_p13_text = (rng.Text or '').rstrip('\r\n\x07')[:30]
                elif page == 14:
                    if y < first_p14_y:
                        first_p14_y = y
                        first_p14_text = (rng.Text or '').rstrip('\r\n\x07')[:30]
                if page > 14:
                    break
            except Exception:
                pass
        print(f'\nLast p13 paragraph: y={last_p13_y} text={last_p13_text!r}')
        print(f'First p14 paragraph: y={first_p14_y} text={first_p14_text!r}')

        # 退職給付引当金繰入 specific lookup
        for pi in range(1, n_paras + 1):
            try:
                para = doc.Paragraphs(pi)
                txt = (para.Range.Text or '').rstrip('\r\n\x07')
                if '退職給付' in txt:
                    rng = doc.Range(para.Range.Start, para.Range.Start)
                    page = int(rng.Information(1))
                    y = float(rng.Information(6))
                    print(f'\n退職給付 paragraph: page={page} y={y} text={txt[:40]!r}')
                    break
            except Exception:
                pass
    except Exception:
        traceback.print_exc()
    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        word.Quit()


if __name__ == '__main__':
    main()
