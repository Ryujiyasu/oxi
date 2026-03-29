#!/usr/bin/env python3
"""COM: Get all paragraph Y positions on page 1."""
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

    print("=== All paragraphs on page 1 ===")
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        rng = p.Range
        page = rng.Information(3)
        if page == 1:
            y = rng.Information(6)
            font = rng.Font.Name
            size = rng.Font.Size
            text = rng.Text[:60].replace('\r', '').replace('\n', '')
            print(f"  P{i}: y={y:.2f}, font={font}, size={size}, \"{text}\"")
        elif page >= 2:
            # Show first 3 on page 2
            y = rng.Information(6)
            font = rng.Font.Name
            size = rng.Font.Size
            text = rng.Text[:60].replace('\r', '').replace('\n', '')
            print(f"  P{i} (PAGE 2): y={y:.2f}, font={font}, size={size}, \"{text}\"")
            if page == 2 and i > 62:
                break

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
