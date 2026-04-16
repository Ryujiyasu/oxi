"""Check paragraph indent via COM."""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    p = doc.Paragraphs(10)
    fmt = p.Format
    print(f"LeftIndent: {fmt.LeftIndent:.1f}pt")
    print(f"RightIndent: {fmt.RightIndent:.1f}pt")
    print(f"FirstLineIndent: {fmt.FirstLineIndent:.1f}pt")
    print(f"CharacterUnitLeftIndent: {fmt.CharacterUnitLeftIndent:.1f}")
    print(f"CharacterUnitRightIndent: {fmt.CharacterUnitRightIndent:.1f}")
    print(f"CharacterUnitFirstLineIndent: {fmt.CharacterUnitFirstLineIndent:.1f}")
finally:
    doc.Close(False)
    word.Quit()
