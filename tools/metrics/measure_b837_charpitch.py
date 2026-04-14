"""Measure charPitch and line break positions for b837."""

import win32com.client
import os
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc_path = os.path.abspath("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
    doc = word.Documents.Open(doc_path, ReadOnly=True)
    time.sleep(1)

    sec = doc.Sections(1)
    ps = sec.PageSetup

    print(f"PageWidth: {ps.PageWidth:.2f}pt")
    print(f"Margins L/R: {ps.LeftMargin:.2f} / {ps.RightMargin:.2f}")
    content_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    print(f"Content width: {content_w:.2f}pt")
    print(f"LayoutMode: {ps.LayoutMode}")

    try:
        print(f"CharsLine: {ps.CharsLine}")
        print(f"GridDistanceHorizontal: {ps.GridDistanceHorizontal:.4f}pt")
    except Exception as e:
        print(f"CharsLine/GridDist: {e}")

    # Measure first few paragraph positions and per-char advances
    for pi in [1, 13]:  # P1 and P13 (from spec §11.2.1 measurement)
        if pi > doc.Paragraphs.Count:
            continue
        p = doc.Paragraphs(pi)
        rng = p.Range
        text = rng.Text[:60].replace('\r', '')
        y = doc.Range(rng.Start, rng.Start+1).Information(6)
        print(f"\nP{pi}: y={y:.1f} \"{text}\"")

        # Per-char advance (Information(5) difference)
        advances = []
        for ci in range(rng.Start, min(rng.End - 1, rng.Start + 40)):
            cr1 = doc.Range(ci, ci+1)
            cr2 = doc.Range(ci+1, ci+2)
            x1 = cr1.Information(5)
            x2 = cr2.Information(5)
            ch = cr1.Text
            adv = x2 - x1
            advances.append((ch, adv))

        for ch, adv in advances[:30]:
            print(f"  '{ch}' adv={adv:.1f}")

    doc.Close(False)
    word.Quit()

if __name__ == '__main__':
    measure()
