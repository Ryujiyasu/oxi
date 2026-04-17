"""Measure the first 10 paragraphs on page 5 of 0e7a to identify what's at
y=125px (= ~60pt) — the content Oxi seems to miss.
"""
import win32com.client
import os

docx_path = os.path.abspath(
    r"tools\golden-test\documents\docx\0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    total = doc.Paragraphs.Count
    # Walk paragraphs; print those on page 5 (first ~10)
    count = 0
    for i in range(1, total + 1):
        p = doc.Paragraphs(i)
        r = p.Range
        pg = r.Information(3)
        if pg != 5:
            continue
        y = r.Information(6)
        fmt = p.Format
        # Get font size from first run
        fs = None
        try:
            first_run = r.Words(1)
            fs = first_run.Font.Size
        except Exception:
            pass
        text = r.Text[:50].replace('\r','\\r').replace('\n','\\n')
        kn = fmt.KeepWithNext
        kt = fmt.KeepTogether
        pb = fmt.PageBreakBefore
        sb = fmt.SpaceBefore
        sa = fmt.SpaceAfter
        print(f"P{i:3d} pg={pg} y={y:7.2f} fs={fs} kn={int(kn)} kt={int(kt)} pb={int(pb)} sb={sb:.1f} sa={sa:.1f} [{text[:40]}]")
        count += 1
        if count > 15:
            break
    doc.Close(False)
finally:
    word.Quit()
