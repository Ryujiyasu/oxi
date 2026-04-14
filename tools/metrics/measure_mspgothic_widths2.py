"""Measure MS PGothic proportional widths via Word COM.
Only measures codepoints that actually appear in test documents (1382 chars).
"""
import win32com.client
import os
import json
import time

def main():
    with open('pipeline_data/mspgothic_codepoints.json') as f:
        codepoints = json.load(f)
    print(f"Codepoints to measure: {len(codepoints)}")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Add()
    sec = doc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72

    font_name = "\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af"  # ＭＳ Ｐゴシック
    font_size = 10.5

    results = {}
    batch_size = 30

    for batch_start in range(0, len(codepoints), batch_size):
        batch = codepoints[batch_start:batch_start + batch_size]

        # Clear document
        doc.Content.Delete()

        # Insert pairs: measured_char + reference_char, one per paragraph
        lines = []
        for cp in batch:
            lines.append(chr(cp) + "A")
        full_text = "\r".join(lines)

        rng = doc.Range(0, 0)
        rng.Text = full_text
        rng.Font.Name = font_name
        rng.Font.Size = font_size
        rng.ParagraphFormat.Alignment = 0  # left
        rng.ParagraphFormat.SpaceAfter = 0
        rng.ParagraphFormat.SpaceBefore = 0

        time.sleep(0.3)

        pos = 0
        for i, cp in enumerate(batch):
            try:
                rng1 = doc.Range(pos, pos + 1)
                x1 = rng1.Information(5)
                rng2 = doc.Range(pos + 1, pos + 2)
                x2 = rng2.Information(5)
                advance_pt = x2 - x1
                results[cp] = round(advance_pt * 20) / 20.0  # round to 0.05pt (1tw)
                pos += 3  # char + A + \r
            except Exception as e:
                pos += 3

        if batch_start % 300 == 0:
            print(f"  {batch_start}/{len(codepoints)} measured...")

    doc.Close(False)

    print(f"\nMeasured {len(results)} characters")

    # Compare with GDI table
    with open('crates/oxidocs-core/src/font/data/gdi_width_overrides.json', encoding='utf-8') as f:
        gdi = json.load(f)
    pg14 = gdi.get('MS PGothic', {}).get('14', {})

    mismatches = 0
    total_diff = 0.0
    for cp, com_pt in results.items():
        gdi_px = pg14.get(str(cp))
        if gdi_px is not None:
            gdi_pt = float(gdi_px) * 72.0 / 96.0
            diff = abs(com_pt - gdi_pt)
            if diff > 0.1:
                mismatches += 1
                total_diff += diff
                if mismatches <= 10:
                    print(f"  U+{cp:04X} ({chr(cp)}): COM={com_pt}pt GDI={gdi_pt}pt diff={diff:.2f}")

    print(f"\nMismatches (>0.1pt): {mismatches}/{len(results)}")
    if mismatches > 0:
        print(f"Average mismatch: {total_diff/mismatches:.2f}pt")

    # Save as COM twips override format
    out = {}
    for cp, pt in results.items():
        out[str(cp)] = pt

    out_path = "pipeline_data/mspgothic_com_widths.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"MS PGothic": {"10.5": out}}, f, indent=2)
    print(f"Saved to {out_path}")

    word.Quit()

if __name__ == "__main__":
    main()
