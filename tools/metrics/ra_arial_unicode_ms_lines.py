#!/usr/bin/env python3
"""COM: Count lines per paragraph on pages 1-3 of the contract document."""
import win32com.client
import os, time

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx"))

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    # Count total lines on each page
    for page_num in range(1, 4):
        # Use built-in line counting
        total_lines = 0
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            rng = p.Range
            page = rng.Information(3)  # wdActiveEndPageNumber
            if page == page_num:
                # Count lines in this paragraph on this page
                start_line = rng.Information(10)  # wdFirstCharacterLineNumber
                # Move to end of paragraph
                end_rng = p.Range
                end_rng.Collapse(0)  # wdCollapseEnd
                end_rng.MoveEnd(1, -1)  # back one char to stay in paragraph
                end_line = end_rng.Information(10)
                num_lines = end_line - start_line + 1
                total_lines += num_lines
            elif page > page_num:
                break
        print(f"Page {page_num}: approximately {total_lines} lines")

    # More precise: use wdNumberOfLinesInDocument
    print(f"\nTotal lines in document: {doc.ComputeStatistics(1)}")  # wdStatisticLines

    # Count lines per page by checking each paragraph
    print("\n=== Page 1 last paragraphs ===")
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        page = p.Range.Information(3)
        if page == 1:
            last_p1 = i
        elif page == 2:
            y = p.Range.Information(6)
            text = p.Range.Text[:40].replace('\r', '')
            if i <= last_p1 + 3:
                print(f"  P{i} (page 2): y={y:.2f}, \"{text}\"")
            first_p2 = i if not 'first_p2' in dir() else first_p2
            last_p2 = i
        elif page == 3:
            first_p3_text = p.Range.Text[:40].replace('\r', '')
            print(f"\n  Last para on page 1: P{last_p1}")
            print(f"  First para on page 2: P{first_p2 if 'first_p2' in locals() else '?'}")
            print(f"  Last para on page 2: P{last_p2}")
            print(f"  First para on page 3: P{i}, \"{first_p3_text}\"")
            break

    # Now get text of last 3 paragraphs on page 2
    print("\n=== Last paragraphs on page 2 ===")
    for idx in range(last_p2 - 2, last_p2 + 2):
        if idx < 1:
            continue
        p = doc.Paragraphs(idx)
        page = p.Range.Information(3)
        y = p.Range.Information(6)
        text = p.Range.Text[:80].replace('\r', '')
        font = p.Range.Font.Name
        size = p.Range.Font.Size
        print(f"  P{idx} (page {page}): y={y:.2f}, font={font}, size={size}, \"{text}\"")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
