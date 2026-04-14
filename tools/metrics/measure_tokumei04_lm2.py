"""Measure Y positions in tokumei_08_04 (linesAndChars, pitch=272, margin=851)."""
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
        for i in range(1, min(n + 1, 25)):
            p = doc.Paragraphs(i)
            r = p.Range
            y = r.Information(6)
            fmt = p.Format
            ls = fmt.LineSpacing
            ls_rule = fmt.LineSpacingRule
            font = r.Font
            fsize = font.Size
            text = r.Text[:25].replace('\r', '\\r').replace('\n', '\\n')
            rule_name = {0: "Single", 1: "1.5", 2: "Double", 3: "atLeast", 4: "exact", 5: "multiple"}
            rule = rule_name.get(ls_rule, str(ls_rule))
            print(f"P{i-1}: y={y:.2f}pt  font={fsize}pt  ls={ls:.2f}({rule})  text=\"{text}\"")
            if i > 1:
                prev_y = doc.Paragraphs(i-1).Range.Information(6)
                gap = y - prev_y
                print(f"     gap: {gap:.2f}pt = {gap*20:.0f}tw")
        doc.Close(False)
    finally:
        word.Quit()

if __name__ == "__main__":
    measure("tools/golden-test/documents/docx/1636d28e2c46_tokumei_08_04.docx")
