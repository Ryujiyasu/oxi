"""Deep measure P1 of outline_08: why is line height 15.5pt instead of 18pt?"""
import win32com.client
import time, os

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "d77a58485f16_20240705_resources_data_outline_08.docx"))

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    p1 = doc.Paragraphs(1)
    r1 = p1.Range
    fmt = p1.Format

    print("=== P1 (empty paragraph) ===")
    print(f"y: {r1.Information(6)}")
    print(f"Font.Name: {r1.Font.Name}")
    print(f"Font.NameFarEast: {r1.Font.NameFarEast}")
    print(f"Font.NameAscii: {r1.Font.NameAscii}")
    print(f"Font.Size: {r1.Font.Size}")
    print(f"LineSpacing: {fmt.LineSpacing}")
    print(f"LineSpacingRule: {fmt.LineSpacingRule}")
    print(f"SpaceBefore: {fmt.SpaceBefore}")
    print(f"SpaceAfter: {fmt.SpaceAfter}")
    print(f"Alignment: {fmt.Alignment}")
    print(f"Style: {p1.Style.NameLocal}")

    # Check LayoutMode
    print(f"\nLayoutMode: {doc.PageSetup.LayoutMode}")
    print(f"LinePitch: {doc.Sections(1).PageSetup.LinePitch}")
    print(f"CharacterPitchAndSpacing: {doc.Sections(1).PageSetup.CharsLine}")

    # P2 for comparison
    p2 = doc.Paragraphs(2)
    r2 = p2.Range
    print(f"\n=== P2 (title) ===")
    print(f"y: {r2.Information(6)}")
    print(f"Font.Name: {r2.Font.Name}")
    print(f"Font.Size: {r2.Font.Size}")
    print(f"LineSpacing: {p2.Format.LineSpacing}")
    print(f"LineSpacingRule: {p2.Format.LineSpacingRule}")
    print(f"Style: {p2.Style.NameLocal}")

    # P3
    p3 = doc.Paragraphs(3)
    r3 = p3.Range
    print(f"\n=== P3 (empty after title) ===")
    print(f"y: {r3.Information(6)}")
    print(f"Font.Name: {r3.Font.Name}")
    print(f"Font.NameFarEast: {r3.Font.NameFarEast}")
    print(f"Font.Size: {r3.Font.Size}")
    print(f"LineSpacing: {p3.Format.LineSpacing}")
    print(f"LineSpacingRule: {p3.Format.LineSpacingRule}")

    # P9 (sz=10.5, gap=17pt)
    p9 = doc.Paragraphs(9)
    r9 = p9.Range
    print(f"\n=== P9 (10.5pt, gap=17pt?) ===")
    print(f"y: {r9.Information(6)}")
    print(f"Font.Name: {r9.Font.Name}")
    print(f"Font.NameFarEast: {r9.Font.NameFarEast}")
    print(f"Font.Size: {r9.Font.Size}")
    print(f"Style: {p9.Style.NameLocal}")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
