"""Measure Word's text Y offset from margin top for grid-snapped documents.

Opens existing .docx files and measures the Y position of the first paragraph.
Compares with margin_top to determine the text_y_offset formula.
"""
import win32com.client
import os, time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    docx_dir = os.path.join(os.path.dirname(__file__), "..", "..", "tools", "golden-test", "documents", "docx")
    docx_dir = os.path.abspath(docx_dir)

    # Documents with known grid settings
    test_docs = [
        "6a39b1812290_001851238.docx",     # lines, pitch=18pt
        "8bc929011c90_001851240.docx",     # lines, pitch=18pt
        "d1e8ac8fd1cc_kyodokenkyuyoushiki06.docx",  # linesAndChars, pitch=14.6pt
        "b837808d0555_20240705_resources_data_guideline_02.docx",  # linesAndChars, pitch=18pt
        "3a4f9fbe1a83_001620506.docx",     # lines, pitch=18pt
        "de6e32b5960b_tokumei_08_01-1.docx",  # linesAndChars, pitch=14.6pt
        "gen2_054_Audit_Report.docx",       # likely no grid
    ]

    print(f"{'Document':50s} {'margin':>7s} {'P1_y':>7s} {'offset':>7s} {'pitch':>6s} {'P1_fs':>6s} {'P1_ls':>6s}")
    print("-" * 95)

    for docx_name in test_docs:
        path = os.path.join(docx_dir, docx_name)
        if not os.path.exists(path):
            print(f"{docx_name}: not found")
            continue

        try:
            doc = word.Documents.Open(path, ReadOnly=True)
            time.sleep(0.5)

            # Margin top
            margin_top = doc.PageSetup.TopMargin

            # Grid pitch
            try:
                grid_pitch = doc.Sections(1).PageSetup.LinePitch
            except:
                grid_pitch = 0

            # First paragraph Y position
            p1 = doc.Paragraphs(1)
            p1_y = p1.Range.Information(6)  # wdVerticalPositionRelativeToPage

            # Font size of first paragraph
            p1_fs = p1.Range.Font.Size

            # Line spacing
            p1_ls = p1.Format.LineSpacing

            offset = p1_y - margin_top

            # Also measure P2 for gap calculation
            if doc.Paragraphs.Count >= 2:
                p2 = doc.Paragraphs(2)
                p2_y = p2.Range.Information(6)
                gap = p2_y - p1_y
            else:
                gap = 0

            short = docx_name[:48]
            print(f"{short:50s} {margin_top:7.2f} {p1_y:7.2f} {offset:7.2f} {grid_pitch:6.1f} {p1_fs:6.1f} {p1_ls:6.1f} gap={gap:.1f}")

            doc.Close(SaveChanges=False)

        except Exception as e:
            print(f"{docx_name}: ERROR {e}")
            try: doc.Close(SaveChanges=False)
            except: pass

    word.Quit()

if __name__ == "__main__":
    measure()
