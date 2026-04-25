"""
COM measurement: measure e3c545 paragraph 30 actual line wrap.

Word shows idx 30 (1-indexed 31) as 1 line via COM lines array,
Oxi wraps to 2 lines. Verify visually by scanning per-char X positions.
"""
import win32com.client
import os

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    docx_path = os.path.abspath(
        "tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx"
    )
    doc = word.Documents.Open(docx_path, ReadOnly=True)
    try:
        # 1-indexed 31 corresponds to 0-indexed 30
        for para_idx in [30, 31, 32]:
            para = doc.Paragraphs(para_idx)
            r = para.Range
            text = r.Text
            print(f"\n=== P{para_idx}: Y={r.Information(6):.2f} page={r.Information(3)} ===")
            print(f"Text length: {len(text)}, first 60 chars: {repr(text[:60])}")
            # Scan per-char Y (group into lines)
            lines = []
            current_line_y = None
            current_line = []
            for i in range(len(text)):
                sub = doc.Range(r.Start + i, r.Start + i + 1)
                try:
                    y = sub.Information(6)
                    x = sub.Information(5)
                except:
                    continue
                if current_line_y is None or abs(y - current_line_y) < 3:
                    current_line_y = y if current_line_y is None else current_line_y
                    current_line.append((i, text[i], x, y))
                else:
                    lines.append(current_line)
                    current_line_y = y
                    current_line = [(i, text[i], x, y)]
            if current_line:
                lines.append(current_line)

            print(f"Physical lines: {len(lines)}")
            for i, ln in enumerate(lines):
                first = ln[0]
                last = ln[-1]
                text_of_line = ''.join(c[1] for c in ln).replace('\r', '\\r').replace('\n', '\\n')
                print(f"  Line {i+1}: y={first[3]:.2f}  start_x={first[2]:.2f}  end_x={last[2]:.2f}  chars={len(ln)}  text={repr(text_of_line[:80])}")

    finally:
        doc.Close(False)
        word.Quit()

if __name__ == "__main__":
    main()
