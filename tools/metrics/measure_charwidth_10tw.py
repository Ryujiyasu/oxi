"""Verify 10tw character width rounding across multiple fonts and sizes.
Creates minimal repro documents and measures x positions via COM."""
import win32com.client
import math

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    test_cases = [
        ("Cambria", 11, "Please note the"),
        ("Cambria", 12, "Please note the"),
        ("Cambria", 10, "Please note the"),
        ("Times New Roman", 11, "Please note the"),
        ("Times New Roman", 12, "Please note the"),
        ("Arial", 11, "Please note the"),
        ("Arial", 10, "Please note the"),
        ("Calibri", 11, "Please note the"),
        ("Calibri", 10.5, "Please note the"),
        ("Century", 11, "Please note the"),
        ("ＭＳ 明朝", 11, "あいうえおかきくけこ"),
        ("ＭＳ 明朝", 10.5, "あいうえおかきくけこ"),
        ("ＭＳ ゴシック", 12, "あいうえおかきくけこ"),
    ]

    for font_name, font_size, test_text in test_cases:
        doc = word.Documents.Add()
        ps = doc.Sections(1).PageSetup
        ps.LayoutMode = 0
        ps.TopMargin = 72
        ps.LeftMargin = 72

        doc.Content.Delete()
        p = doc.Paragraphs(1)
        p.Range.Text = test_text
        p.Range.Font.Name = font_name
        p.Range.Font.Size = font_size
        p.Format.SpaceBefore = 0
        p.Format.SpaceAfter = 0
        p.Format.LineSpacingRule = 0  # Single

        start = p.Range.Start
        chars = list(test_text)

        print(f"=== {font_name} {font_size}pt ===")
        mismatches_1tw = 0
        mismatches_10tw = 0
        total_chars = 0

        for i in range(min(len(chars), 15)):
            r = doc.Range(start + i, start + i + 1)
            x = r.Information(5)
            if i > 0:
                r_prev = doc.Range(start + i - 1, start + i)
                x_prev = r_prev.Information(5)
                word_advance = x - x_prev

                # We need the advance width from font metrics
                # For now just show the word advance
                ch = chars[i-1]
                total_chars += 1

                # Check if word_advance matches 10tw or 1tw pattern
                word_tw = word_advance * 20
                is_10tw = abs(word_tw - round(word_tw / 10) * 10) < 0.5
                is_1tw = abs(word_tw - round(word_tw)) < 0.5

                if not is_1tw:
                    mismatches_1tw += 1
                if not is_10tw:
                    mismatches_10tw += 1

                if i <= 10:
                    print(f"  {ch}: advance={word_advance:.2f}pt = {word_tw:.1f}tw "
                          f"{'10tw' if is_10tw else '1tw' if is_1tw else '??'}")

        if total_chars > 0:
            print(f"  Summary: {total_chars} chars, "
                  f"10tw match={total_chars-mismatches_10tw}/{total_chars}, "
                  f"1tw match={total_chars-mismatches_1tw}/{total_chars}")
        print()

        doc.Close(False)

finally:
    word.Quit()
