"""Measure b837 Page 1 line-by-line break points (Word COM ground truth).

Prints each paragraph on page 1 with: line index, start char, end char,
char count, and last 4 chars — to compare against Oxi rendering output.
"""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    n_paras = doc.Paragraphs.Count
    print(f"Total paragraphs: {n_paras}")
    for pi in range(1, n_paras + 1):
        para = doc.Paragraphs(pi)
        rng = para.Range
        # Skip paragraphs not on page 1
        first_char = rng.Characters(1)
        try:
            page = first_char.Information(3)  # wdActiveEndPageNumber
        except Exception:
            continue
        if page != 1:
            if page > 1:
                break
            continue
        text = rng.Text.rstrip('\r\n\x07')
        if not text.strip():
            continue
        # Collect char-by-char line numbers to detect break points
        lines = {}
        n_chars = len(text)
        for ci in range(n_chars):
            try:
                c = rng.Characters(ci + 1)
                ln = c.Information(10)  # wdFirstCharacterLineNumber
                lines.setdefault(ln, []).append((ci, text[ci]))
            except Exception:
                break
        print(f"\n--- Para {pi} (page {page}, {n_chars} chars) ---")
        print(f"  head: {text[:30]!r}")
        for ln in sorted(lines.keys()):
            chars = lines[ln]
            line_text = ''.join(c for _, c in chars)
            first_i = chars[0][0]
            last_i = chars[-1][0]
            tail = line_text[-6:] if len(line_text) > 6 else line_text
            print(f"  L{ln}: [{first_i:3d}..{last_i:3d}] ({len(chars):2d} chars) tail={tail!r}")
finally:
    doc.Close(False)
    word.Quit()
