"""Measure the actual page bottom position for db9ca18368cd document."""

import win32com.client
import os
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc_path = os.path.abspath("tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx")
    doc = word.Documents.Open(doc_path, ReadOnly=True)
    time.sleep(1)

    sec = doc.Sections(1)
    ps = sec.PageSetup
    print(f"Page height: {ps.PageHeight:.2f}pt")
    print(f"Top margin: {ps.TopMargin:.2f}pt")
    print(f"Bottom margin: {ps.BottomMargin:.2f}pt")
    print(f"Header dist: {ps.HeaderDistance:.2f}pt")
    print(f"Footer dist: {ps.FooterDistance:.2f}pt")
    print(f"Content height: {ps.PageHeight - ps.TopMargin - ps.BottomMargin:.2f}pt")
    print(f"Page bottom: {ps.PageHeight - ps.BottomMargin:.2f}pt")

    # Find the last text on each page
    for page_num in range(1, 4):
        print(f"\n--- Page {page_num} ---")
        # Find last paragraph on this page
        for i in range(doc.Paragraphs.Count, 0, -1):
            p = doc.Paragraphs(i)
            rng = p.Range
            # Check first char of paragraph
            first_char = doc.Range(rng.Start, rng.Start + 1)
            pg = first_char.Information(3)
            if pg == page_num:
                y = first_char.Information(6)
                text = rng.Text[:50].replace('\r', '')
                print(f"  Last para starting on page {page_num}: P{i} y={y:.1f} [{len(rng.Text)-1}c] {text[:40]}")

                # Find the last char on this page
                for ci in range(rng.Start, min(rng.End, rng.Start + 500)):
                    cr = doc.Range(ci, ci + 1)
                    cp = cr.Information(3)
                    if cp != page_num:
                        # Previous char was last on this page
                        prev = doc.Range(ci - 1, ci)
                        print(f"  Last char on page {page_num}: offset={ci-1-rng.Start} y={prev.Information(6):.1f} char={repr(prev.Text)}")
                        break
                break

    # Also check footer position
    try:
        footer = sec.Footers(1)  # wdHeaderFooterPrimary
        if footer.Range.Text.strip():
            print(f"\nFooter text: {footer.Range.Text[:30]}")
            print(f"Footer y: {footer.Range.Information(6):.1f}")
    except:
        pass

    doc.Close(False)
    word.Quit()

if __name__ == '__main__':
    measure()
