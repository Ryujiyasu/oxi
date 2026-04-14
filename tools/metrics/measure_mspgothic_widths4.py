"""Measure MS PGothic proportional widths correctly.
Uses same-char repetition (e.g. ああああ) to avoid autoSpaceDE contamination.
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

    font_name = "\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af"
    font_size = 10.5

    # Strategy: insert each char repeated 3 times on its own paragraph
    # Measure advance as x[1] - x[0] (no autoSpaceDE between same chars)
    valid_cps = []
    lines = []
    for cp in codepoints:
        try:
            ch = chr(cp)
            ch.encode('cp932')
            lines.append(ch * 3)  # repeat 3 times
            valid_cps.append(cp)
        except:
            pass

    print(f"Valid: {len(valid_cps)} chars")

    full_text = "\r".join(lines)
    rng = doc.Range(0, 0)
    rng.Text = full_text
    rng.Font.Name = font_name
    rng.Font.Size = font_size
    rng.ParagraphFormat.Alignment = 0  # left
    rng.ParagraphFormat.SpaceAfter = 0
    rng.ParagraphFormat.SpaceBefore = 0

    print("Document created, measuring...")
    time.sleep(3)

    results = {}
    pos = 0
    errors = 0
    for i, cp in enumerate(valid_cps):
        try:
            r0 = doc.Range(pos, pos + 1)
            r1 = doc.Range(pos + 1, pos + 2)
            x0 = r0.Information(5)
            x1 = r1.Information(5)
            advance = x1 - x0
            if advance > 0 and advance < 30:
                # Convert to twips (round to nearest twip)
                tw = round(advance * 20)
                results[cp] = tw
            pos += 4  # 3 chars + \r
        except:
            errors += 1
            pos += 4

        if i % 200 == 0 and i > 0:
            print(f"  {i}/{len(valid_cps)} ({errors} errors)...")

    # Save BEFORE closing Word
    out = {}
    for cp, tw in results.items():
        out[str(cp)] = float(tw)

    out_path = "pipeline_data/mspgothic_com_widths_v2.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"MS PGothic": {"10.5": out}}, f, indent=2)
    print(f"\nSaved {len(results)} measurements to {out_path}")

    # Sample
    for cp in [0x3001, 0x3042, 0x30A2, 0x4E00, 0x7B2C, 0xFF1A]:
        if cp in results:
            tw = results[cp]
            print(f"  U+{cp:04X} ({chr(cp)}): {tw}tw = {tw/20:.2f}pt")

    # Compare with GDI
    with open('crates/oxidocs-core/src/font/data/gdi_width_overrides.json', encoding='utf-8') as f:
        gdi = json.load(f)
    pg14 = gdi.get('MS PGothic', {}).get('14', {})

    mismatches = 0
    for cp, tw in results.items():
        gdi_px = pg14.get(str(cp))
        if gdi_px is not None:
            gdi_tw = round(float(gdi_px) * 72 / 96 * 20)
            if abs(tw - gdi_tw) > 2:
                mismatches += 1

    print(f"\nMismatches (>0.1pt): {mismatches}/{len(results)}")

    try:
        doc.Close(False)
        word.Quit()
    except:
        pass

if __name__ == "__main__":
    main()
