"""
Ra: 0e7af1ae8f21 p.2 の行ごとのY座標・テキスト・行幅をCOM計測
Oxiとの差分を特定する
"""
import os
import json
import win32com.client
import pythoncom

DOCX = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx"
))
OUT = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..",
    "pipeline_data", "com_measurements", "0e7af1ae8f21_p2_lines.json"
))


def main():
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3

    try:
        doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
        try:
            results = {"paragraphs": []}

            # Page setup
            sec = doc.Sections(1).PageSetup
            results["page_setup"] = {
                "page_width": round(sec.PageWidth, 2),
                "page_height": round(sec.PageHeight, 2),
                "margin_top": round(sec.TopMargin, 2),
                "margin_bottom": round(sec.BottomMargin, 2),
                "margin_left": round(sec.LeftMargin, 2),
                "margin_right": round(sec.RightMargin, 2),
            }
            content_w = sec.PageWidth - sec.LeftMargin - sec.RightMargin
            results["page_setup"]["content_width"] = round(content_w, 2)

            # Find paragraphs on p.2 and nearby
            rules = {0: "single", 1: "1.5", 2: "double", 3: "atLeast", 4: "exactly", 5: "multiple"}

            for i in range(1, doc.Paragraphs.Count + 1):
                para = doc.Paragraphs(i)
                rng = para.Range
                page = rng.Information(3)
                if page < 2:
                    continue
                if page > 2:
                    break

                y = rng.Information(6)
                fmt = para.Format
                ls = fmt.LineSpacing
                lr = rules.get(fmt.LineSpacingRule, str(fmt.LineSpacingRule))
                indent_l = fmt.LeftIndent
                indent_r = fmt.RightIndent
                indent_first = fmt.FirstLineIndent
                sa = fmt.SpaceAfter
                sb = fmt.SpaceBefore

                font_name = rng.Font.Name
                font_size = rng.Font.Size
                text = rng.Text.replace("\r", "").replace("\n", "")

                # Count lines in this paragraph
                # Use range line counting
                start = rng.Start
                end = rng.End
                # Move to start of para, then count lines
                rng2 = doc.Range(start, start)
                start_line = rng2.Information(10)  # wdFirstCharacterLineNumber
                rng3 = doc.Range(end - 1, end - 1) if end > start else rng2
                end_line = rng3.Information(10)
                line_count = end_line - start_line + 1 if end > start else 1

                results["paragraphs"].append({
                    "idx": i,
                    "y_pt": round(y, 2),
                    "text": text[:60],
                    "text_len": len(text),
                    "line_count": line_count,
                    "line_spacing": round(ls, 2),
                    "line_rule": lr,
                    "indent_left": round(indent_l, 2),
                    "indent_right": round(indent_r, 2),
                    "indent_first": round(indent_first, 2),
                    "space_before": round(sb, 2),
                    "space_after": round(sa, 2),
                    "font": font_name,
                    "font_size": font_size,
                })

            # Print summary
            ps = results["page_setup"]
            print(f"Page: {ps['page_width']}x{ps['page_height']}pt")
            print(f"Margins: T={ps['margin_top']} B={ps['margin_bottom']} L={ps['margin_left']} R={ps['margin_right']}")
            print(f"Content width: {ps['content_width']}pt")
            print()

            for p in results["paragraphs"]:
                lines = p["line_count"]
                il = f" iL={p['indent_left']}" if p['indent_left'] else ""
                ir = f" iR={p['indent_right']}" if p['indent_right'] else ""
                i1 = f" i1={p['indent_first']}" if p['indent_first'] else ""
                print(f"P{p['idx']} y={p['y_pt']} {lines}L {p['font']} {p['font_size']}pt{il}{ir}{i1} [{p['text_len']}ch] \"{p['text'][:50]}\"")

            # Save
            os.makedirs(os.path.dirname(OUT), exist_ok=True)
            with open(OUT, "w", encoding="utf-8") as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            print(f"\nSaved to {OUT}")

        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
