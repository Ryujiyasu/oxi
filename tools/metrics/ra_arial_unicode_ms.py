#!/usr/bin/env python3
"""COM measurement: Arial Unicode MS document - font resolution and line heights."""
import win32com.client
import os, json, time

DOCX = os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx")
DOCX = os.path.abspath(DOCX)

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    results = {}

    # 1. Check actual font used (what Word resolves Arial Unicode MS to)
    print("=== Font Resolution ===")
    for i in range(1, min(21, doc.Paragraphs.Count + 1)):
        p = doc.Paragraphs(i)
        rng = p.Range
        font_name = rng.Font.Name
        font_size = rng.Font.Size
        text = rng.Text[:40].replace('\r', '').replace('\n', '')
        print(f"  P{i}: font={font_name}, size={font_size}pt, text=\"{text}\"")

    # 2. Measure Y positions for page 2 paragraphs
    print("\n=== Page 2 Paragraph Y Positions ===")
    page2_paras = []
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        rng = p.Range
        page_num = rng.Information(3)  # wdActiveEndPageNumber
        if page_num == 2:
            y_pos = rng.Information(6)  # wdVerticalPositionRelativeToPage
            font_name = rng.Font.Name
            font_size = rng.Font.Size
            text = rng.Text[:50].replace('\r', '').replace('\n', '')
            page2_paras.append({
                "para_idx": i,
                "y_pt": round(y_pos, 4),
                "font": font_name,
                "size": font_size,
                "text": text
            })
            print(f"  P{i}: y={y_pos:.2f}pt, font={font_name}, size={font_size}pt, \"{text}\"")
        elif page_num == 3:
            # Get first para of page 3 to know where page 2 ends
            y_pos = rng.Information(6)
            text = rng.Text[:50].replace('\r', '').replace('\n', '')
            page2_paras.append({
                "para_idx": i,
                "y_pt": round(y_pos, 4),
                "font": rng.Font.Name,
                "size": rng.Font.Size,
                "text": text,
                "note": "first_para_page3"
            })
            print(f"  P{i} (PAGE 3 START): y={y_pos:.2f}pt, \"{text}\"")
            break

    # 3. Compute line height gaps between consecutive paragraphs on page 2
    print("\n=== Line Height Gaps (page 2) ===")
    for j in range(1, len(page2_paras)):
        prev = page2_paras[j-1]
        curr = page2_paras[j]
        gap = curr["y_pt"] - prev["y_pt"]
        print(f"  P{prev['para_idx']}→P{curr['para_idx']}: gap={gap:.4f}pt")

    # 4. Check document grid settings
    print("\n=== Section/Grid Settings ===")
    for i in range(1, doc.Sections.Count + 1):
        sec = doc.Sections(i)
        pf = sec.PageSetup
        print(f"  Section {i}: linePitch={pf.LinePitch}pt, topMargin={pf.TopMargin}pt, bottomMargin={pf.BottomMargin}pt")
        print(f"    pageHeight={pf.PageHeight}pt, pageWidth={pf.PageWidth}pt")
        print(f"    headerDist={pf.HeaderDistance}pt, footerDist={pf.FooterDistance}pt")

    # 5. Check last paragraph on page 2 vs first on page 3
    print("\n=== Page boundary ===")
    last_p2 = None
    first_p3 = None
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        page_num = p.Range.Information(3)
        if page_num == 2:
            last_p2 = i
        elif page_num == 3 and first_p3 is None:
            first_p3 = i
            break
    print(f"  Last para on page 2: P{last_p2}")
    print(f"  First para on page 3: P{first_p3}")
    if last_p2:
        rng = doc.Paragraphs(last_p2).Range
        y = rng.Information(6)
        text = rng.Text[:60].replace('\r', '').replace('\n', '')
        print(f"  P{last_p2}: y={y:.2f}pt, \"{text}\"")

    results["page2_paras"] = page2_paras

    # Save
    out_path = os.path.join(os.path.dirname(__file__), "..", "..",
        "pipeline_data", "com_measurements", "arial_unicode_ms_p2.json")
    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {out_path}")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
