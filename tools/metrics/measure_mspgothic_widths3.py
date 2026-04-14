"""Measure MS PGothic proportional widths via Word COM.
Uses one document with all characters at once (no repeated doc creation).
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
    time.sleep(2)
    sec = doc.Sections(1)
    ps = sec.PageSetup
    ps.LeftMargin = 72
    ps.RightMargin = 72
    ps.TopMargin = 36
    ps.BottomMargin = 36

    font_name = "\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af"
    font_size = 10.5

    # Insert ALL characters at once: each char + 'A' on a separate paragraph
    lines = []
    valid_cps = []
    for cp in codepoints:
        try:
            chr(cp).encode('cp932')
            lines.append(chr(cp) + "A")
            valid_cps.append(cp)
        except:
            pass

    print(f"Valid (CP932 encodable): {len(valid_cps)}")

    # Insert in one shot
    full_text = "\r".join(lines)
    rng = doc.Range(0, 0)
    rng.Text = full_text
    rng.Font.Name = font_name
    rng.Font.Size = font_size
    rng.ParagraphFormat.Alignment = 0  # left
    rng.ParagraphFormat.SpaceAfter = 0
    rng.ParagraphFormat.SpaceBefore = 0
    rng.ParagraphFormat.LineSpacingRule = 0  # single

    print("Document created, waiting for layout...")
    time.sleep(3)

    results = {}
    pos = 0
    errors = 0
    for i, cp in enumerate(valid_cps):
        try:
            rng1 = doc.Range(pos, pos + 1)
            x1 = rng1.Information(5)
            rng2 = doc.Range(pos + 1, pos + 2)
            x2 = rng2.Information(5)
            advance_pt = x2 - x1
            if advance_pt > 0:
                results[cp] = round(advance_pt * 20) / 20.0
            pos += 3  # char + A + \r
        except Exception as e:
            errors += 1
            pos += 3

        if i % 200 == 0 and i > 0:
            print(f"  {i}/{len(valid_cps)} measured ({errors} errors)...")

    # Save results BEFORE closing (Word may crash on Close)
    out = {}
    for cp, pt in results.items():
        out[str(cp)] = pt
    out_path = "pipeline_data/mspgothic_com_widths.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"MS PGothic": {"10.5": out}}, f, indent=2)
    print(f"Saved {len(results)} measurements to {out_path}")

    try:
        doc.Close(False)
    except:
        pass

    print(f"\nMeasured {len(results)} characters ({errors} errors)")

    # Sample results
    sample_cps = [0x3001, 0x3002, 0x3042, 0x30A2, 0x4E00, 0x7B2C, 0x672C, 0xFF08, 0xFF1A]
    for cp in sample_cps:
        if cp in results:
            print(f"  U+{cp:04X} ({chr(cp)}): {results[cp]}pt")

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

    print(f"\nMismatches (>0.1pt): {mismatches}/{len(results)}")
    if mismatches > 0:
        print(f"Average mismatch: {total_diff/mismatches:.2f}pt")

    # Save
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
