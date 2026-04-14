"""Measure Y positions in b837 (linesAndChars, pitch=360, margin=1021) via Word COM."""
import win32com.client
import os, time

def measure(docx_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path))
        time.sleep(1)
        n = doc.Paragraphs.Count
        print(f"Total paragraphs: {n}")
        for i in range(1, min(n + 1, 20)):
            p = doc.Paragraphs(i)
            r = p.Range
            y = r.Information(6)
            fmt = p.Format
            ls = fmt.LineSpacing
            ls_rule = fmt.LineSpacingRule
            font = r.Font
            fsize = font.Size
            text = r.Text[:30].replace('\r', '\\r').replace('\n', '\\n')
            rule_name = {0: "Single", 1: "1.5", 2: "Double", 3: "atLeast", 4: "exact", 5: "multiple"}
            rule = rule_name.get(ls_rule, str(ls_rule))
            # Count lines in this paragraph
            lines = r.Information(14)  # wdNumberOfLinesInDocument... not right
            print(f"P{i-1}: y={y:.2f}pt  font={fsize}pt  ls={ls:.2f}({rule})  text=\"{text}\"")
            if i > 1:
                prev_y = doc.Paragraphs(i-1).Range.Information(6)
                gap = y - prev_y
                print(f"     gap: {gap:.2f}pt = {gap*20:.0f}tw")
        doc.Close(False)
    finally:
        word.Quit()

if __name__ == "__main__":
    measure("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
