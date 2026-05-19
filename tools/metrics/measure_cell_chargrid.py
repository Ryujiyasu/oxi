"""Measure Word's per-char x positions inside cell repros via PDF export + pymupdf.

This avoids the COM Information(5) unreliability seen in earlier hanging measurements.
Same approach as measure_lla_word.py.
"""
import json, os, sys, time
import win32com.client
import pymupdf

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPRO_DIR = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "tools/metrics/cell_chargrid_repro")
PDF_DIR = os.path.abspath(REPRO_DIR + "_pdf")
OUT = os.path.abspath(sys.argv[2] if len(sys.argv) > 2 else "tools/metrics/cell_chargrid_word.json")
os.makedirs(PDF_DIR, exist_ok=True)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

results = {}
for fname in sorted(os.listdir(REPRO_DIR)):
    if not fname.endswith(".docx"):
        continue
    label = fname[:-5]
    docx = os.path.join(REPRO_DIR, fname)
    pdf_path = os.path.join(PDF_DIR, f"{label}.pdf")
    print(f"\n=== {label} ===")
    if not os.path.exists(pdf_path):
        doc = word.Documents.Open(docx, ReadOnly=True)
        time.sleep(0.3)
        doc.ExportAsFixedFormat(pdf_path, 17)  # wdExportFormatPDF=17
        doc.Close(SaveChanges=False)
    pdf = pymupdf.open(pdf_path)
    page = pdf[0]
    text_dict = page.get_text("dict")
    cell_data = []
    for blk in text_dict.get("blocks", []):
        for ln in blk.get("lines", []):
            for span in ln.get("spans", []):
                text = span["text"]
                if any(c in text for c in "組織的管理措置"):
                    bb = span["bbox"]
                    cell_data.append({
                        "text": text,
                        "x_start": round(bb[0], 3),
                        "x_end": round(bb[2], 3),
                        "y": round(bb[1], 3),
                        "fs": round(span["size"], 2),
                        "n_chars": len(text),
                        "width": round(bb[2] - bb[0], 3),
                        "per_char": round((bb[2] - bb[0]) / max(len(text), 1), 3),
                    })
    pdf.close()
    for c in cell_data:
        print(f"  text={c['text']!r} fs={c['fs']} N={c['n_chars']} W={c['width']:.2f}pt per_char={c['per_char']:.3f}pt y={c['y']:.2f}")
    results[label] = cell_data

word.Quit()

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(f"\nSaved to {OUT}")
