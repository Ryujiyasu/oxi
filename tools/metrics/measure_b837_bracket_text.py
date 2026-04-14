"""Measure per-character advances for b837 date block (inside bracket shape)."""
import win32com.client, os, time, sys

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc_path = os.path.abspath("tools/golden-test/documents/docx/b837808d0555_20240705_resources_data_guideline_02.docx")
    doc = word.Documents.Open(doc_path, ReadOnly=True)
    time.sleep(1)

    # P3-P9 (COM 1-indexed) are the date block lines inside bracket
    # Bracket: margin-left=244.1pt, width=213.8pt → right edge=457.9pt
    # margin-top=17.95pt
    for pi in range(3, 10):
        p = doc.Paragraphs(pi)
        rng = p.Range
        text = rng.Text.replace('\r', '')
        y = doc.Range(rng.Start, rng.Start+1).Information(6)
        x_start = doc.Range(rng.Start, rng.Start+1).Information(5)

        sys.stdout.buffer.write(f"\nP{pi}: y={y:.1f} x_start={x_start:.1f} \"{text}\"\n".encode('utf-8'))

        # Per-char advance
        advances = []
        for ci in range(rng.Start, rng.End - 1):
            c1 = doc.Range(ci, ci+1)
            c2 = doc.Range(ci+1, ci+2)
            x1 = c1.Information(5)
            x2 = c2.Information(5)
            ch = c1.Text
            adv = x2 - x1
            advances.append((ch, x1, adv))

        for ch, x, adv in advances:
            sys.stdout.buffer.write(f"  '{ch}' x={x:.1f} adv={adv:.1f}\n".encode('utf-8'))

        # Total width
        if advances:
            total = sum(a for _, _, a in advances)
            last_x = advances[-1][1] + advances[-1][2]
            sys.stdout.buffer.write(f"  total_adv={total:.1f} last_x={last_x:.1f}\n".encode('utf-8'))

    doc.Close(False)
    word.Quit()

if __name__ == '__main__':
    measure()
