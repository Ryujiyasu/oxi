"""Measure b837 line heights and Y positions."""
import win32com.client, os, time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc_path = os.path.abspath("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
    doc = word.Documents.Open(doc_path, ReadOnly=True)
    time.sleep(1)

    # Page 1 paragraph positions
    print("Page 1 paragraphs:")
    for i in range(1, min(doc.Paragraphs.Count+1, 25)):
        p = doc.Paragraphs(i)
        rng = p.Range
        sc = doc.Range(rng.Start, rng.Start+1)
        y = sc.Information(6)
        pg = sc.Information(3)
        if pg > 1: break
        cc = len(rng.Text) - 1
        txt = rng.Text[:30].replace('\r','')
        print(f"  P{i}: y={y:.1f} [{cc}c] \"{txt}\"")

    # Line-level for P3 (long paragraph)
    p3 = doc.Paragraphs(3)
    r3 = p3.Range
    print(f"\nP3 lines:")
    prev_y = None
    for ci in range(r3.Start, min(r3.End, r3.Start+500)):
        cr = doc.Range(ci, ci+1)
        cy = cr.Information(6)
        if prev_y is None or abs(cy-prev_y)>1:
            print(f"  y={cy:.1f} off={ci-r3.Start}")
            prev_y = cy

    doc.Close(False)
    word.Quit()

if __name__ == '__main__':
    measure()
