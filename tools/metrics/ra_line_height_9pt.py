"""
Ra: MS明朝 9pt の正確な行高さをCOM計測で確定
複数フォント・複数サイズでの行高さを系統的に計測し、計算式を特定する
"""
import os
import json
import win32com.client
import pythoncom

OUT = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..",
    "pipeline_data", "com_measurements", "line_height_systematic.json"
))


def main():
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3

    results = []

    try:
        # Create test doc with NO grid (like the contract doc)
        doc = word.Documents.Add()
        sec = doc.Sections(1).PageSetup
        sec.LayoutMode = 0  # wdLayoutModeDefault (no grid)

        fonts_and_sizes = [
            ("ＭＳ 明朝", [8, 8.5, 9, 9.5, 10, 10.5, 11, 12, 14, 16]),
            ("ＭＳ ゴシック", [8, 8.5, 9, 9.5, 10, 10.5, 11, 12, 14, 16]),
            ("游明朝", [8, 9, 10, 10.5, 11, 12, 14]),
            ("游ゴシック", [8, 9, 10, 10.5, 11, 12, 14]),
            ("メイリオ", [8, 9, 10, 10.5, 11, 12]),
            ("Calibri", [8, 9, 10, 10.5, 11, 12]),
            ("Century", [8, 9, 10, 10.5, 11, 12]),
            ("Arial", [8, 9, 10, 10.5, 11, 12]),
        ]

        for font_name, sizes in fonts_and_sizes:
            for size in sizes:
                # Clear doc
                doc.Content.Text = ""

                # Add 5 identical paragraphs with this font/size
                for _ in range(5):
                    rng = doc.Content
                    rng.Collapse(0)
                    rng.Text = "あいうえおかきくけこ漢字テスト ABCDEF\r"
                    rng.Font.Name = font_name
                    rng.Font.Size = size

                # Measure Y positions
                ys = []
                pages = []
                for i in range(1, min(doc.Paragraphs.Count + 1, 6)):
                    p = doc.Paragraphs(i)
                    y = p.Range.Information(6)
                    pg = p.Range.Information(3)
                    ys.append(round(y, 4))
                    pages.append(pg)

                # Calculate line height from consecutive same-page paragraphs
                diffs = []
                for i in range(len(ys) - 1):
                    if pages[i] == pages[i + 1]:
                        diffs.append(round(ys[i + 1] - ys[i], 4))

                if diffs:
                    avg = sum(diffs) / len(diffs)
                    entry = {
                        "font": font_name,
                        "size": size,
                        "ppem": round(size * 96 / 72),
                        "ys": ys,
                        "diffs": diffs,
                        "line_height": round(avg, 4),
                    }
                    results.append(entry)
                    print(f"{font_name} {size}pt (ppem={entry['ppem']}): lh={avg:.4f}pt  diffs={diffs}")

        doc.Close(SaveChanges=False)

        # Save
        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved to {OUT}")

    finally:
        word.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
