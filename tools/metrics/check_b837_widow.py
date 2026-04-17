"""Check widow/orphan control for b837 p#49 and related paragraphs."""
import win32com.client, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

DOC = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\b837808d0555_20240705_resources_data_guideline_02.docx"
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    for idx in [39, 49, 60]:
        p = wdoc.Paragraphs(idx)
        f = p.Format
        print(f"p#{idx}: WidowControl={f.WidowControl}  KeepWithNext={f.KeepWithNext}  KeepTogether={f.KeepTogether}  PageBreakBefore={f.PageBreakBefore}")
        print(f"       text={p.Range.Text.strip()[:50]}")
finally:
    wdoc.Close(False)
    word.Quit()
