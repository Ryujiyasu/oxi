"""
Measure comprehensive_test.docx - v2 with better style detection.
"""
import win32com.client
import os, time

DOCX_PATH = os.path.abspath(r"tests/fixtures/comprehensive_test.docx")

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Open(DOCX_PATH)
        time.sleep(2)
        doc.Repaginate()
        time.sleep(1)

        n_paras = doc.Paragraphs.Count

        # Get all para details
        print("=== Paragraph Details ===")
        for i in range(1, min(n_paras + 1, 40)):
            para = doc.Paragraphs(i)
            rng = para.Range
            text = rng.Text.strip()[:50].encode('ascii', 'replace').decode('ascii')

            rng_start = para.Range
            rng_start.Collapse(1)
            y = rng_start.Information(6)

            fmt = para.Format
            style = para.Style

            # Get style outline level to detect headings
            outline_level = fmt.OutlineLevel  # 1-9 for headings, 10 for body

            line_spacing = fmt.LineSpacing
            line_rule = fmt.LineSpacingRule
            space_before = fmt.SpaceBefore
            space_after = fmt.SpaceAfter

            # Line spacing rules: 0=AtLeast, 1=Exactly, 2=wdLineSpaceDouble,
            # 3=wdLineSpace1pt5, 4=wdLineSpaceSingle, 5=wdLineSpaceMultiple
            rule_names = {0: "AtLeast", 1: "Exactly", 2: "Double", 3: "1.5", 4: "Single", 5: "Multiple"}
            rule_name = rule_names.get(line_rule, str(line_rule))

            heading = f" [H{outline_level}]" if outline_level <= 9 else ""

            # Font info
            font_name = rng.Font.Name
            font_size = rng.Font.Size

            print(f"P{i:2d} Y={y:7.2f} Before={space_before:5.1f} After={space_after:5.1f} "
                  f"LS={line_spacing:6.2f}({rule_name:>8}) "
                  f"Font={font_name}:{font_size}pt{heading} "
                  f"| {text}")

        # Specific investigation: mixed font sizes paragraph
        print("\n=== Mixed Font Sizes (Para 4) ===")
        para4 = doc.Paragraphs(4)
        rng4 = para4.Range
        n_chars = rng4.Characters.Count
        print(f"Total characters: {n_chars}")

        # Check each run's font size
        words = rng4.Words
        n_words = words.Count
        for w in range(1, min(n_words + 1, 20)):
            wd = words(w)
            txt = wd.Text.encode('ascii', 'replace').decode('ascii')
            print(f"  Word {w}: '{txt}' size={wd.Font.Size}pt font={wd.Font.Name}")

        # Check grid snap settings per paragraph
        print("\n=== Snap to Grid Check ===")
        for i in range(1, min(n_paras + 1, 15)):
            para = doc.Paragraphs(i)
            snap = para.Format.SnapToGrid
            text = para.Range.Text.strip()[:40].encode('ascii', 'replace').decode('ascii')
            print(f"P{i:2d}: SnapToGrid={snap} | {text}")

        # Get document grid pitch
        print("\n=== Section Grid ===")
        sec = doc.Sections(1)
        pf = sec.PageSetup
        print(f"LinesPage={pf.LinesPage}")

        # Calculate grid pitch = (pageH - topM - botM) / linesPage
        page_h = pf.PageHeight
        top_m = pf.TopMargin
        bot_m = pf.BottomMargin
        lines = pf.LinesPage
        pitch = (page_h - top_m - bot_m) / lines if lines > 0 else 0
        print(f"PageHeight={page_h}, TopMargin={top_m}, BottomMargin={bot_m}")
        print(f"Calculated grid pitch: {pitch:.4f}pt ({pitch*20:.2f}tw)")

        # Also check default paragraph style
        print("\n=== Style Investigation ===")
        for i in range(1, min(doc.Styles.Count + 1, 10)):
            try:
                s = doc.Styles(i)
                name = s.NameLocal.encode('ascii', 'replace').decode('ascii')
                print(f"Style {i}: '{name}' (built-in={s.BuiltIn})")
            except:
                pass

        doc.Close(0)
    finally:
        word.Quit()

if __name__ == "__main__":
    main()
