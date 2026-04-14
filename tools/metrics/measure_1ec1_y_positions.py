"""Measure Y positions of all paragraphs in 1ec1 document via Word COM."""
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
        print()

        for i in range(1, min(n + 1, 30)):
            p = doc.Paragraphs(i)
            r = p.Range
            y = r.Information(6)  # wdVerticalPositionRelativeToPage
            x = r.Information(5)  # wdHorizontalPositionRelativeToPage

            # Get formatting
            fmt = p.Format
            ls = fmt.LineSpacing
            ls_rule = fmt.LineSpacingRule
            sb = fmt.SpaceBefore
            sa = fmt.SpaceAfter

            # Get font info
            font = r.Font
            fname = font.Name
            fsize = font.Size

            text = r.Text[:40].replace('\r', '\\r').replace('\n', '\\n')

            rule_name = {0: "Single", 1: "1.5", 2: "Double", 3: "atLeast", 4: "exact", 5: "multiple"}
            rule = rule_name.get(ls_rule, str(ls_rule))

            print(f"P{i-1}: y={y:.2f}pt  x={x:.2f}pt  font={fname}/{fsize}pt  "
                  f"ls={ls:.2f}({rule})  sb={sb:.2f}  sa={sa:.2f}  "
                  f"text=\"{text}\"")

            if i > 1:
                prev_y = doc.Paragraphs(i-1).Range.Information(6)
                gap = y - prev_y
                print(f"     gap from P{i-2}: {gap:.2f}pt")

        doc.Close(False)
    finally:
        word.Quit()

if __name__ == "__main__":
    measure("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
