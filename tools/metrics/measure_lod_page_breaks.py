"""COM: Measure where LOD_Handbook breaks pages."""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
path = os.path.abspath("tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(1)

prev_page = 0
for pi in range(1, doc.Paragraphs.Count + 1):
    para = doc.Paragraphs(pi)
    rng = para.Range
    page = rng.Information(3)
    y = rng.Information(6)
    text = rng.Text.rstrip('\r')

    if page != prev_page:
        if prev_page > 0:
            print(f"  --- Page break {prev_page} -> {page} ---")
        print(f"Page {page}:")
        prev_page = page

    # Only print first/last para per page and page break boundaries
    if pi <= 3 or (page != prev_page) or (text and len(text) > 0):
        pass  # Print all non-empty

    if len(text) > 0 or True:
        is_empty = len(text) == 0
        if pi < 5 or page != prev_page or is_empty or y > 700:
            print(f"  P{pi}: y={y:.1f} pg={page} {'(empty)' if is_empty else text[:30]}")

doc.Close(SaveChanges=False)
word.Quit()
