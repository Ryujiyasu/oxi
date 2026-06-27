# -*- coding: utf-8 -*-
"""Regenerate the corpus word_png references with the CURRENT Word, into a NEW
dir (word_png_new) so the originals are untouched until verified+swapped.

The stored references are 3 months old (2026-03-20) and INCONSISTENT (proven:
byte-identical-template gen2 docs have word_pngs preferring opposite line-height
models — different Word renders). A consistent single-session regeneration is the
sanctioned fix (see [[gen2_vertical_drift]]). Robust subprocess-per-doc with
timeout + kill-hung-Word, mirroring pipeline/word_renderer.py.

Usage: python tools/metrics/regen_word_png.py [start_idx] [limit]
"""
import os, sys, time, subprocess
from pathlib import Path
sys.path.insert(0, r"c:\Users\ryuji\oxi-main")
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
sys.stdout.reconfigure(encoding="utf-8")

REPO = r"c:\Users\ryuji\oxi-main"
DOCS = Path(REPO) / "tools" / "golden-test" / "documents" / "docx"
NEW = Path(REPO) / "pipeline_data" / "word_png_new"
TIMEOUT = 60

RENDER_SCRIPT = r'''
import sys, os
docx_path, out_dir, dpi = sys.argv[1], sys.argv[2], int(sys.argv[3])
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
word.AutomationSecurity = 3
try:
    word.Options.UpdateLinksAtOpen = False
except Exception:
    pass
try:
    doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
    try:
        page_count = doc.ComputeStatistics(2)
        for page_num in range(1, page_count + 1):
            pdf_path = os.path.join(out_dir, f"page_{page_num:04d}.pdf")
            png_path = os.path.join(out_dir, f"page_{page_num:04d}.png")
            doc.ExportAsFixedFormat(OutputFileName=pdf_path, ExportFormat=17,
                OpenAfterExport=False, OptimizeFor=0, Range=3, From=page_num, To=page_num)
            import fitz
            d = fitz.open(pdf_path)
            zoom = dpi / 72
            d[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom)).save(png_path)
            d.close()
            os.unlink(pdf_path)
    finally:
        doc.Close(SaveChanges=False)
finally:
    word.Quit()
    pythoncom.CoUninitialize()
'''


def kill_word():
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True, timeout=10)
    time.sleep(0.5)


def main():
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    limit = int(sys.argv[2]) if len(sys.argv) > 2 else 0
    wp = Path(WORD_PNG_DIR)
    docs = sorted(d.name for d in wp.iterdir()
                  if d.is_dir() and (d / "page_0001.png").exists()
                  and (DOCS / (d.name + ".docx")).exists())
    if limit:
        docs = docs[start:start + limit]
    else:
        docs = docs[start:]
    NEW.mkdir(parents=True, exist_ok=True)
    kill_word()
    ok = fail = skip = 0
    for i, doc_id in enumerate(docs):
        out_dir = NEW / doc_id
        if out_dir.exists() and (out_dir / "page_0001.png").exists():
            skip += 1
            continue
        out_dir.mkdir(parents=True, exist_ok=True)
        docx = str(DOCS / (doc_id + ".docx"))
        try:
            r = subprocess.run([sys.executable, "-c", RENDER_SCRIPT, os.path.abspath(docx),
                                str(out_dir), str(RENDER_DPI)],
                               capture_output=True, text=True, encoding="utf-8",
                               errors="replace", timeout=TIMEOUT)
            n = len(list(out_dir.glob("page_*.png")))
            if r.returncode == 0 and n > 0:
                ok += 1
            else:
                fail += 1
                sys.stderr.write(f"[NG] {doc_id}: rc={r.returncode} pages={n} {r.stderr[:120]}\n")
        except subprocess.TimeoutExpired:
            fail += 1
            sys.stderr.write(f"[TIMEOUT] {doc_id}\n")
            kill_word()
        except Exception as e:
            fail += 1
            sys.stderr.write(f"[ERR] {doc_id}: {e}\n")
        if (i + 1) % 10 == 0:
            sys.stderr.write(f"  {i+1}/{len(docs)}  ok={ok} fail={fail} skip={skip}\n")
            sys.stderr.flush()
    print(f"DONE: ok={ok} fail={fail} skip={skip} / {len(docs)} docs")


if __name__ == "__main__":
    main()
