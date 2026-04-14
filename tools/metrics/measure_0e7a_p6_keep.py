"""Check keepNext/keepTogether/widowOrphan for paragraphs around p6/p7 boundary."""
import win32com.client
import os

docx_path = os.path.abspath(r"tools\golden-test\documents\docx\0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(docx_path, ReadOnly=True)

    for i in range(225, 240):
        para = doc.Paragraphs(i)
        rng = para.Range
        page = rng.Information(3)
        y = rng.Information(6)
        fmt = para.Format

        keep_next = fmt.KeepWithNext
        keep_together = fmt.KeepTogether
        widow = fmt.WidowControl
        outline_level = fmt.OutlineLevel  # 10=body text
        text = rng.Text[:50].replace('\r', '\\r').replace('\n', '\\n')

        print(f"P{i:3d} page={page} y={y:7.1f} keepNext={keep_next} keepTogether={keep_together} widow={widow} outline={outline_level} [{text[:35]}]")

    doc.Close(False)
finally:
    word.Quit()
